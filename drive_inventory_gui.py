#!/usr/bin/env python3
"""Interface PySide6 do DriveAutomate."""

from __future__ import annotations

import datetime as dt
import sys
import threading
from pathlib import Path
from typing import Any, Callable

from PySide6.QtCore import (
    QObject,
    QSettings,
    QStandardPaths,
    QThread,
    QTimer,
    QUrl,
    Qt,
    Signal,
    Slot,
)
from PySide6.QtGui import QColor, QCloseEvent, QDesktopServices, QFont, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from drive_downloader import (
    DownloadCancelled,
    DownloadStats,
    InventoryWorkbookError,
    WorkbookScan,
    download_from_inventory,
    scan_inventory_workbook,
)
from drive_inventory import (
    ExportCancelled,
    authenticate,
    authenticate_new_account,
    build_drive_service,
    default_accounts_dir,
    default_token_path,
    export_drive_inventory,
    format_user_error,
    legacy_default_token_path,
    resolve_drive_source,
    token_identity,
)
from oauth_config import (
    has_client_config,
    install_client_config,
    oauth_client_path,
)

APP_TITLE = "DriveAutomate"
APP_VERSION = "3.0.0"


def default_downloads_dir() -> Path:
    location = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
    return Path(location) if location else Path.home() / "Downloads"


def default_inventory_path() -> Path:
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return default_downloads_dir() / f"inventario_google_drive_{timestamp}.xlsx"


def open_local_path(path: str | Path) -> bool:
    target = Path(path).resolve()
    if not target.exists():
        raise FileNotFoundError(f"Arquivo ou pasta não encontrado: {target}")
    return QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))


class OperationWorker(QObject):
    """Executa uma função bloqueante fora da thread da interface."""

    log = Signal(str)
    progress = Signal(int, int, int)
    succeeded = Signal(object)
    failed = Signal(str)
    cancelled = Signal(str)
    finished = Signal()

    def __init__(self, operation: Callable[["OperationWorker"], Any]) -> None:
        super().__init__()
        self.operation = operation
        self.cancel_event = threading.Event()

    def request_cancel(self) -> None:
        self.cancel_event.set()

    @Slot()
    def run(self) -> None:
        try:
            result = self.operation(self)
            if self.cancel_event.is_set():
                self.cancelled.emit("Operação cancelada pelo usuário.")
            else:
                self.succeeded.emit(result)
        except (ExportCancelled, DownloadCancelled) as exc:
            self.cancelled.emit(str(exc))
        except Exception as exc:  # mensagem será apresentada na thread principal
            if isinstance(exc, (InventoryWorkbookError, RuntimeError)):
                message = str(exc)
            else:
                message = format_user_error(exc)
            self.failed.emit(message)
        finally:
            self.finished.emit()


class DriveAutomateWindow(QMainWindow):
    """Janela principal voltada a usuários não técnicos."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE} {APP_VERSION}")
        # A área de conteúdo precisa de espaço para os campos e os botões em
        # telas de 100% de escala. O tamanho mínimo evita que os controles se
        # sobreponham quando a janela é redimensionada.
        self.resize(1120, 820)
        self.setMinimumSize(980, 720)

        self.settings = QSettings("DriveAutomate", "DriveAutomate")
        self.accounts: list[Path] = []
        self.active_token: Path | None = None
        self.thread: QThread | None = None
        self.worker: OperationWorker | None = None
        self.operation_result: object | None = None
        self.operation_error: str | None = None
        self.operation_cancelled: str | None = None
        self.success_handler: Callable[[object], None] | None = None
        self.operation_cancellable = True
        self.pending_close = False
        self.action_widgets: list[QWidget] = []

        self._build_ui()
        self._build_menu()
        self._apply_style()
        self.refresh_accounts()
        self._refresh_oauth_status()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        header = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel("DriveAutomate")
        title.setObjectName("appTitle")
        subtitle = QLabel("Inventarie e baixe arquivos do Google Drive com segurança")
        subtitle.setObjectName("subtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        header.addLayout(title_box)
        header.addStretch()
        version = QLabel(f"v{APP_VERSION}")
        version.setObjectName("versionBadge")
        header.addWidget(version)
        layout.addLayout(header)

        account_card = QFrame()
        account_card.setObjectName("card")
        account_layout = QGridLayout(account_card)
        account_layout.setContentsMargins(16, 14, 16, 14)
        account_layout.setHorizontalSpacing(10)
        account_layout.setVerticalSpacing(8)
        account_layout.addWidget(QLabel("Conta Google ativa"), 0, 0)
        self.account_combo = QComboBox()
        self.account_combo.setMinimumWidth(260)
        self.account_combo.setMinimumHeight(36)
        self.account_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.account_combo.currentIndexChanged.connect(self._account_changed)
        # A conta ocupa uma linha inteira; os três botões ficam abaixo. Isso
        # evita que a grade comprima a combo e faça os controles se cobrirem
        # em janelas menores.
        account_layout.addWidget(self.account_combo, 0, 1, 1, 4)
        self.add_account_button = QPushButton("Adicionar conta")
        self.add_account_button.setMinimumWidth(138)
        self.add_account_button.clicked.connect(self.start_add_account)
        self.remove_account_button = QPushButton("Remover conta")
        self.remove_account_button.setMinimumWidth(126)
        self.remove_account_button.clicked.connect(self.remove_account)
        self.oauth_button = QPushButton("Configurar OAuth")
        self.oauth_button.setMinimumWidth(138)
        self.oauth_button.clicked.connect(self.configure_oauth)
        account_actions = QHBoxLayout()
        account_actions.setSpacing(10)
        account_actions.addStretch()
        account_actions.addWidget(self.add_account_button)
        account_actions.addWidget(self.remove_account_button)
        account_actions.addWidget(self.oauth_button)
        account_layout.addLayout(account_actions, 1, 0, 1, 5)
        self.account_status = QLabel()
        self.account_status.setObjectName("muted")
        account_layout.addWidget(self.account_status, 2, 0, 1, 3)
        self.oauth_status = QLabel()
        self.oauth_status.setObjectName("muted")
        account_layout.addWidget(self.oauth_status, 2, 3, 1, 2)
        account_layout.setColumnStretch(1, 1)
        layout.addWidget(account_card)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_inventory_tab(), "1. Inventário")
        self.tabs.addTab(self._build_download_tab(), "2. Downloads")
        layout.addWidget(self.tabs, 1)

        status_bar = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        status_bar.addWidget(self.progress, 1)
        self.cancel_button = QPushButton("Cancelar")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_operation)
        status_bar.addWidget(self.cancel_button)
        layout.addLayout(status_bar)

        self.status_label = QLabel("Pronto para começar.")
        self.status_label.setObjectName("statusLabel")
        layout.addWidget(self.status_label)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(2_000)
        self.log_view.setPlaceholderText("O andamento aparecerá aqui.")
        self.log_view.setMaximumHeight(145)
        layout.addWidget(self.log_view)

        self.action_widgets.extend(
            [
                self.account_combo,
                self.add_account_button,
                self.remove_account_button,
                self.oauth_button,
            ]
        )

    def _build_menu(self) -> None:
        help_menu = self.menuBar().addMenu("Ajuda")
        about_action = help_menu.addAction("Sobre e privacidade")
        about_action.triggered.connect(self.show_about)
        licenses_action = help_menu.addAction("Licenças de terceiros")
        licenses_action.triggered.connect(self.show_licenses)

    def show_about(self) -> None:
        QMessageBox.information(
            self,
            f"Sobre o {APP_TITLE}",
            f"{APP_TITLE} {APP_VERSION}\n\n"
            "Software livre sob licença MIT. Interface construída com PySide6/Qt.\n\n"
            "O aplicativo não possui telemetria, não envia planilhas a terceiros e "
            "armazena OAuth e tokens somente no perfil deste usuário; os tokens são "
            "protegidos pelo DPAPI do Windows. O acesso ao Google Drive é somente leitura.",
        )

    def show_licenses(self) -> None:
        QMessageBox.information(
            self,
            "Licenças de terceiros",
            "PySide6 / Qt for Python: edição comunitária LGPLv3/GPLv3 ou licença "
            "comercial da Qt.\n\nGoogle API Client: Apache-2.0; openpyxl: MIT; "
            "Requests: Apache-2.0.\n\nConsulte THIRD_PARTY_NOTICES.md no pacote da release.",
        )

    def _build_inventory_tab(self) -> QWidget:
        tab = QWidget()
        tab.setMinimumHeight(390)
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)

        intro = QLabel(
            "Cole o link de uma pasta do Meu Drive, compartilhada com você ou de um "
            "Drive compartilhado. O aplicativo detecta a origem automaticamente."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        form = QFormLayout()
        form.setSpacing(12)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.folder_input = QLineEdit()
        self.folder_input.setMinimumWidth(460)
        self.folder_input.setMinimumHeight(36)
        self.folder_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.folder_input.setPlaceholderText("https://drive.google.com/drive/folders/...")
        form.addRow("Pasta do Google Drive", self.folder_input)

        output_container = QWidget()
        output_container.setMinimumHeight(50)
        output_row = QHBoxLayout(output_container)
        output_row.setContentsMargins(0, 0, 0, 0)
        self.inventory_output = QLineEdit(str(default_inventory_path()))
        self.inventory_output.setMinimumWidth(460)
        self.inventory_output.setMinimumHeight(36)
        self.inventory_output.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        output_row.addWidget(self.inventory_output, 1)
        choose_output = QPushButton("Escolher…")
        choose_output.setMinimumWidth(112)
        choose_output.clicked.connect(self.choose_inventory_output)
        output_row.addWidget(choose_output)
        form.addRow("Salvar inventário em", output_container)
        layout.addLayout(form)

        hint = QLabel(
            "O Excel será criado com a coluna amarela “Baixar?”. Ela começa em “Não” "
            "para evitar downloads acidentais."
        )
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QHBoxLayout()
        buttons.addStretch()
        self.test_access_button = QPushButton("Testar acesso")
        self.test_access_button.clicked.connect(self.start_access_test)
        buttons.addWidget(self.test_access_button)
        self.export_button = QPushButton("Gerar inventário Excel")
        self.export_button.setObjectName("primaryButton")
        self.export_button.clicked.connect(self.start_export)
        buttons.addWidget(self.export_button)
        layout.addLayout(buttons)
        layout.addStretch()

        self.action_widgets.extend(
            [self.folder_input, self.inventory_output, choose_output, self.test_access_button, self.export_button]
        )
        return self._scrollable_tab(tab)

    def _build_download_tab(self) -> QWidget:
        tab = QWidget()
        tab.setMinimumHeight(430)
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)

        intro = QLabel(
            "Selecione um inventário gerado pelo DriveAutomate. Você pode marcar “Sim” "
            "na coluna Baixar? ou apagar linhas e usar o modo alternativo."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        form = QFormLayout()
        form.setSpacing(12)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        excel_container = QWidget()
        excel_container.setMinimumHeight(50)
        excel_row = QHBoxLayout(excel_container)
        excel_row.setContentsMargins(0, 0, 0, 0)
        self.download_workbook = QLineEdit()
        self.download_workbook.setMinimumWidth(460)
        self.download_workbook.setMinimumHeight(36)
        self.download_workbook.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.download_workbook.setPlaceholderText("Selecione o inventário .xlsx")
        excel_row.addWidget(self.download_workbook, 1)
        choose_excel = QPushButton("Escolher…")
        choose_excel.setMinimumWidth(112)
        choose_excel.clicked.connect(self.choose_download_workbook)
        excel_row.addWidget(choose_excel)
        form.addRow("Planilha de seleção", excel_container)

        folder_container = QWidget()
        folder_container.setMinimumHeight(50)
        folder_row = QHBoxLayout(folder_container)
        folder_row.setContentsMargins(0, 0, 0, 0)
        self.download_destination = QLineEdit(str(default_downloads_dir() / "DriveAutomate"))
        self.download_destination.setMinimumWidth(460)
        self.download_destination.setMinimumHeight(36)
        self.download_destination.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        folder_row.addWidget(self.download_destination, 1)
        choose_folder = QPushButton("Escolher…")
        choose_folder.setMinimumWidth(112)
        choose_folder.clicked.connect(self.choose_download_destination)
        folder_row.addWidget(choose_folder)
        form.addRow("Salvar arquivos em", folder_container)

        self.selection_mode = QComboBox()
        self.selection_mode.setMinimumHeight(36)
        self.selection_mode.addItem("Somente linhas marcadas como Sim (recomendado)", True)
        self.selection_mode.addItem("Todos os arquivos restantes na planilha", False)
        form.addRow("Modo de seleção", self.selection_mode)

        self.skip_existing = QCheckBox(
            "Pular arquivos existentes"
        )
        self.skip_existing.setMinimumHeight(32)
        self.skip_existing.setToolTip(
            "Desmarcado: arquivos com o mesmo nome serão salvos com um novo nome."
        )
        self.skip_existing.setChecked(True)
        form.addRow("Conflitos", self.skip_existing)
        layout.addLayout(form)

        conflict_hint = QLabel("Desmarque para criar outro nome quando já existir um arquivo.")
        conflict_hint.setObjectName("muted")
        conflict_hint.setWordWrap(True)
        layout.addWidget(conflict_hint)

        warning = QLabel(
            "Arquivos Google Workspace são convertidos automaticamente. Arquivos grandes "
            "podem passar por uma etapa de preparação no Google antes do download."
        )
        warning.setObjectName("hint")
        warning.setWordWrap(True)
        layout.addWidget(warning)

        buttons = QHBoxLayout()
        buttons.addStretch()
        self.scan_button = QPushButton("Analisar planilha")
        self.scan_button.clicked.connect(self.start_scan)
        buttons.addWidget(self.scan_button)
        self.download_button = QPushButton("Baixar selecionados")
        self.download_button.setObjectName("primaryButton")
        self.download_button.clicked.connect(self.start_download_confirmation)
        buttons.addWidget(self.download_button)
        layout.addLayout(buttons)
        layout.addStretch()

        self.action_widgets.extend(
            [
                self.download_workbook,
                self.download_destination,
                choose_excel,
                choose_folder,
                self.selection_mode,
                self.skip_existing,
                self.scan_button,
                self.download_button,
            ]
        )
        return self._scrollable_tab(tab)

    @staticmethod
    def _scrollable_tab(content: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setObjectName("tabScrollArea")
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(content)
        return scroll

    def _apply_style(self) -> None:
        app = QApplication.instance()
        if app is not None:
            app.setStyle("Fusion")
            palette = QPalette()
            palette.setColor(QPalette.ColorRole.Window, QColor("#111827"))
            palette.setColor(QPalette.ColorRole.WindowText, QColor("#E5E7EB"))
            palette.setColor(QPalette.ColorRole.Base, QColor("#111827"))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#1F2937"))
            palette.setColor(QPalette.ColorRole.Text, QColor("#E5E7EB"))
            palette.setColor(QPalette.ColorRole.Button, QColor("#374151"))
            palette.setColor(QPalette.ColorRole.ButtonText, QColor("#F3F4F6"))
            palette.setColor(QPalette.ColorRole.Highlight, QColor("#155E75"))
            palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#F8FAFC"))
            palette.setColor(QPalette.ColorRole.ToolTipBase, QColor("#1F2937"))
            palette.setColor(QPalette.ColorRole.ToolTipText, QColor("#F3F4F6"))
            app.setPalette(palette)
            self.setPalette(palette)
        self.setFont(QFont("Segoe UI", 10))
        self.setStyleSheet(
            """
            QMainWindow, QWidget { background: #111827; color: #E5E7EB; }
            QLabel#appTitle { font-size: 26px; font-weight: 700; color: #F8FAFC; }
            QLabel#subtitle, QLabel#muted { color: #94A3B8; }
            QLabel#versionBadge { background: #1E3A5F; color: #93C5FD; padding: 5px 9px; border-radius: 9px; font-weight: 600; }
            QLabel#hint { background: #243447; color: #BAE6FD; padding: 10px; border: 1px solid #155E75; border-radius: 6px; }
            QLabel#statusLabel { color: #CBD5E1; font-weight: 600; }
            QFrame#card { background: #1F2937; border: 1px solid #374151; border-radius: 10px; }
            QTabWidget::pane { background: #1F2937; border: 1px solid #374151; border-radius: 8px; top: -1px; }
            QTabBar::tab { background: #111827; color: #94A3B8; padding: 11px 18px; border: 1px solid #374151; border-bottom: none; }
            QTabBar::tab:selected { background: #1F2937; color: #67E8F9; font-weight: 700; }
            QLineEdit, QComboBox, QPlainTextEdit { background: #111827; color: #E5E7EB; border: 1px solid #4B5563; border-radius: 6px; padding: 7px; selection-background-color: #155E75; }
            QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus { border: 1px solid #22D3EE; }
            QComboBox QAbstractItemView { background: #1F2937; color: #E5E7EB; selection-background-color: #155E75; selection-color: #F8FAFC; border: 1px solid #4B5563; }
            QPushButton { background: #374151; color: #F3F4F6; border: 1px solid #4B5563; border-radius: 6px; min-height: 34px; padding: 7px 14px; }
            QPushButton:hover { background: #4B5563; }
            QPushButton:pressed { background: #155E75; }
            QPushButton:disabled { color: #6B7280; background: #1F2937; border-color: #374151; }
            QPushButton#primaryButton { background: #0F766E; color: #F0FDFA; border: 1px solid #14B8A6; font-weight: 700; }
            QPushButton#primaryButton:hover { background: #0D9488; }
            QCheckBox { spacing: 8px; color: #E5E7EB; }
            QCheckBox::indicator { width: 16px; height: 16px; }
            QProgressBar { border: 1px solid #4B5563; border-radius: 5px; text-align: center; color: #E5E7EB; background: #111827; }
            QProgressBar::chunk { background: #0F766E; border-radius: 4px; }
            QToolTip { background: #1F2937; color: #F3F4F6; border: 1px solid #4B5563; padding: 5px; }
            """
        )

    def log(self, message: str) -> None:
        self.log_view.appendPlainText(f"{dt.datetime.now():%H:%M:%S}  {message}")

    def _refresh_oauth_status(self) -> None:
        configured = has_client_config()
        self.oauth_status.setText(
            "OAuth configurado neste computador."
            if configured
            else "OAuth ainda não configurado."
        )

    def configure_oauth(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar credencial OAuth desktop",
            str(default_downloads_dir()),
            "Arquivos JSON (*.json)",
        )
        if not filename:
            return
        try:
            install_client_config(filename)
        except Exception as exc:
            QMessageBox.critical(self, "OAuth inválido", str(exc))
            return
        self._refresh_oauth_status()
        self.log("Cliente OAuth configurado no perfil local.")
        QMessageBox.information(
            self,
            "OAuth configurado",
            "Configuração salva somente neste computador. Agora você pode adicionar uma conta.",
        )

    def refresh_accounts(self, selected: Path | None = None) -> None:
        accounts_dir = default_accounts_dir()
        accounts_dir.mkdir(parents=True, exist_ok=True)
        paths: list[Path] = []
        legacy_paths = (legacy_default_token_path(), default_token_path())
        for legacy in legacy_paths:
            if legacy.exists():
                paths.append(legacy)
        for pattern in ("*.dat", "*.json"):
            paths.extend(
                path for path in sorted(accounts_dir.glob(pattern)) if path not in paths
            )
        self.accounts = paths

        previous = selected or self.active_token
        self.account_combo.blockSignals(True)
        self.account_combo.clear()
        for path in paths:
            email, display_name = token_identity(path)
            if email:
                label = f"{display_name} ({email})" if display_name else email
            elif path in legacy_paths:
                label = "Conta legada"
            else:
                label = f"Conta local {path.stem[:8]}"
            self.account_combo.addItem(label, str(path))
        if previous in paths:
            self.account_combo.setCurrentIndex(paths.index(previous))
        elif paths:
            saved = self.settings.value("active_account", "")
            index = next((i for i, path in enumerate(paths) if str(path) == saved), 0)
            self.account_combo.setCurrentIndex(index)
        self.account_combo.blockSignals(False)
        self._account_changed(self.account_combo.currentIndex())

    @Slot(int)
    def _account_changed(self, index: int) -> None:
        self.active_token = self.accounts[index] if 0 <= index < len(self.accounts) else None
        if self.active_token:
            self.settings.setValue("active_account", str(self.active_token))
            self.account_status.setText(f"Conta selecionada: {self.account_combo.currentText()}")
        else:
            self.account_status.setText("Nenhuma conta conectada.")

    def _require_account(self) -> Path | None:
        if not has_client_config():
            QMessageBox.warning(
                self,
                "Configure o OAuth",
                "Clique em 'Configurar OAuth' e importe a credencial desktop antes de continuar.",
            )
            return None
        if not self.active_token:
            QMessageBox.warning(
                self,
                "Conecte uma conta",
                "Use 'Adicionar conta' e selecione a conta que possui acesso aos arquivos.",
            )
            return None
        return self.active_token

    def start_add_account(self) -> None:
        if not has_client_config():
            self._require_account()
            return

        def operation(worker: OperationWorker):
            return authenticate_new_account(str(oauth_client_path()), None)

        def success(result: object) -> None:
            token_path, email, display_name = result  # type: ignore[misc]
            self.refresh_accounts(Path(token_path))
            QMessageBox.information(
                self, "Conta conectada", f"Conta conectada: {display_name or email}"
            )

        self.start_operation(
            "Conclua a autorização no navegador…",
            operation,
            success,
            cancellable=False,
        )

    def remove_account(self) -> None:
        if not self.active_token or not self.active_token.exists():
            QMessageBox.information(self, "Contas", "Nenhuma conta está selecionada.")
            return
        answer = QMessageBox.question(
            self,
            "Remover conta",
            "Remover esta conta somente deste computador? Os arquivos do Drive não serão alterados.",
        )
        if answer != QMessageBox.Yes:
            return
        label = self.account_combo.currentText()
        self.active_token.unlink()
        self.active_token = None
        self.refresh_accounts()
        self.log(f"Conta local removida: {label}")

    def choose_inventory_output(self) -> None:
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar inventário",
            self.inventory_output.text() or str(default_inventory_path()),
            "Planilha Excel (*.xlsx)",
        )
        if filename:
            if not filename.casefold().endswith(".xlsx"):
                filename += ".xlsx"
            self.inventory_output.setText(filename)

    def choose_download_workbook(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar inventário",
            str(default_downloads_dir()),
            "Planilha Excel (*.xlsx)",
        )
        if filename:
            self.download_workbook.setText(filename)

    def choose_download_destination(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Selecionar pasta para downloads",
            self.download_destination.text() or str(default_downloads_dir()),
        )
        if folder:
            self.download_destination.setText(folder)

    def start_access_test(self) -> None:
        token = self._require_account()
        folder = self.folder_input.text().strip()
        if not token or not folder:
            if token and not folder:
                QMessageBox.warning(self, "Informe a pasta", "Cole o link ou ID da pasta.")
            return

        def operation(worker: OperationWorker):
            creds = authenticate(str(oauth_client_path()), token)
            service = build_drive_service(creds)
            return resolve_drive_source(service, folder, True)

        def success(result: object) -> None:
            source = result
            label = source.source_type  # type: ignore[attr-defined]
            if source.drive_name:  # type: ignore[attr-defined]
                label += f" — {source.drive_name}"  # type: ignore[attr-defined]
            QMessageBox.information(
                self,
                "Acesso confirmado",
                f"Pasta: {source.root_name}\nOrigem: {label}",  # type: ignore[attr-defined]
            )

        self.start_operation("Testando acesso à pasta…", operation, success)

    def start_export(self) -> None:
        token = self._require_account()
        folder = self.folder_input.text().strip()
        output = self.inventory_output.text().strip()
        if not token:
            return
        if not folder:
            QMessageBox.warning(self, "Informe a pasta", "Cole o link ou ID da pasta.")
            return
        if not output.casefold().endswith(".xlsx"):
            QMessageBox.warning(self, "Destino inválido", "Escolha um arquivo com extensão .xlsx.")
            return

        def operation(worker: OperationWorker):
            return export_drive_inventory(
                folder,
                output,
                include_shared_drives=True,
                credentials_path=str(oauth_client_path()),
                token_path=token,
                progress_callback=worker.log.emit,
                cancel_callback=worker.cancel_event.is_set,
            )

        def success(result: object) -> None:
            count = int(result)
            self.download_workbook.setText(output)
            answer = QMessageBox.question(
                self,
                "Inventário concluído",
                f"{count:,} item(ns) exportado(s). Deseja abrir o Excel agora?",
            )
            if answer == QMessageBox.Yes:
                open_local_path(output)

        self.start_operation("Lendo o Google Drive e gerando o Excel…", operation, success)

    def _download_inputs(self) -> tuple[str, str, bool] | None:
        workbook = self.download_workbook.text().strip()
        destination = self.download_destination.text().strip()
        if not workbook:
            QMessageBox.warning(self, "Selecione a planilha", "Escolha o inventário Excel.")
            return None
        if not destination:
            QMessageBox.warning(self, "Selecione o destino", "Escolha onde salvar os arquivos.")
            return None
        return workbook, destination, bool(self.selection_mode.currentData())

    def start_scan(self) -> None:
        inputs = self._download_inputs()
        if not inputs:
            return
        workbook, _destination, selected_only = inputs

        def operation(worker: OperationWorker):
            return scan_inventory_workbook(workbook, selected_only=selected_only)

        def success(result: object) -> None:
            scan: WorkbookScan = result  # type: ignore[assignment]
            QMessageBox.information(self, "Análise concluída", scan.message())

        self.start_operation("Analisando a planilha…", operation, success)

    def start_download_confirmation(self) -> None:
        token = self._require_account()
        inputs = self._download_inputs()
        if not token or not inputs:
            return
        workbook, destination, selected_only = inputs

        def operation(worker: OperationWorker):
            return scan_inventory_workbook(workbook, selected_only=selected_only)

        def success(result: object) -> None:
            scan: WorkbookScan = result  # type: ignore[assignment]
            if scan.candidates == 0:
                QMessageBox.information(self, "Nada selecionado", scan.message())
                return
            warning = ""
            if not selected_only:
                warning = "\n\nVocê escolheu baixar TODAS as linhas restantes."
            answer = QMessageBox.question(
                self,
                "Confirmar downloads",
                scan.message() + warning + "\n\nDeseja continuar?",
            )
            if answer == QMessageBox.Yes:
                QTimer.singleShot(
                    0,
                    lambda: self._start_download(
                        token, workbook, destination, selected_only
                    ),
                )

        self.start_operation("Validando a seleção…", operation, success)

    def _start_download(
        self,
        token: Path,
        workbook: str,
        destination: str,
        selected_only: bool,
    ) -> None:
        skip_existing = self.skip_existing.isChecked()

        def operation(worker: OperationWorker):
            creds = authenticate(str(oauth_client_path()), token)
            return download_from_inventory(
                workbook,
                destination,
                creds,
                selected_only=selected_only,
                skip_existing=skip_existing,
                log_callback=worker.log.emit,
                progress_callback=worker.progress.emit,
                cancel_callback=worker.cancel_event.is_set,
            )

        def success(result: object) -> None:
            stats: DownloadStats = result  # type: ignore[assignment]
            message = (
                f"Concluídos: {stats.downloaded:,}\n"
                f"Já existentes: {stats.skipped_existing:,}\n"
                f"Renomeados por conflito: {stats.renamed_existing:,}\n"
                f"Não suportados: {stats.skipped_unsupported:,}\n"
                f"Falhas: {stats.failed:,}"
            )
            answer = QMessageBox.question(
                self,
                "Downloads concluídos",
                message + "\n\nDeseja abrir a pasta de destino?",
            )
            if answer == QMessageBox.Yes:
                open_local_path(destination)

        self.start_operation("Baixando os arquivos selecionados…", operation, success)

    def start_operation(
        self,
        status: str,
        operation: Callable[[OperationWorker], object],
        success_handler: Callable[[object], None] | None = None,
        *,
        cancellable: bool = True,
    ) -> None:
        if self.thread and self.thread.isRunning():
            return
        self.operation_result = None
        self.operation_error = None
        self.operation_cancelled = None
        self.success_handler = success_handler
        self.operation_cancellable = cancellable
        self.status_label.setText(status)
        self.progress.setRange(0, 0)
        self.progress.setValue(0)
        self._set_busy(True, cancellable)

        thread = QThread(self)
        worker = OperationWorker(operation)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.log.connect(self.log)
        worker.progress.connect(self._update_progress)
        worker.succeeded.connect(self._store_result)
        worker.failed.connect(self._store_error)
        worker.cancelled.connect(self._store_cancelled)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._finish_operation)
        self.thread = thread
        self.worker = worker
        thread.start()

    def _set_busy(self, busy: bool, cancellable: bool = True) -> None:
        for widget in self.action_widgets:
            widget.setEnabled(not busy)
        self.cancel_button.setEnabled(busy and cancellable)

    @Slot(int, int, int)
    def _update_progress(self, current: int, total: int, file_percentage: int) -> None:
        if total <= 0:
            self.progress.setRange(0, 0)
            return
        self.progress.setRange(0, 1000)
        overall = ((max(current - 1, 0) + file_percentage / 100) / total) * 1000
        self.progress.setValue(min(int(overall), 1000))
        self.progress.setFormat(f"{current:,}/{total:,} — arquivo {file_percentage}%")

    @Slot(object)
    def _store_result(self, result: object) -> None:
        self.operation_result = result

    @Slot(str)
    def _store_error(self, message: str) -> None:
        self.operation_error = message

    @Slot(str)
    def _store_cancelled(self, message: str) -> None:
        self.operation_cancelled = message

    @Slot()
    def _finish_operation(self) -> None:
        result = self.operation_result
        error = self.operation_error
        cancelled = self.operation_cancelled
        handler = self.success_handler
        self.thread = None
        self.worker = None
        self.success_handler = None
        self.operation_cancellable = True
        self._set_busy(False)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("%p%")

        if error:
            self.status_label.setText("Não foi possível concluir.")
            self.log(f"ERRO: {error}")
            QMessageBox.critical(self, "Atenção", error)
        elif cancelled:
            self.status_label.setText("Operação cancelada.")
            self.log(cancelled)
        else:
            self.status_label.setText("Operação concluída.")
            if handler:
                handler(result)

        if self.pending_close:
            self.pending_close = False
            QTimer.singleShot(0, self.close)

    def cancel_operation(self) -> None:
        if not self.worker:
            return
        self.worker.request_cancel()
        self.cancel_button.setEnabled(False)
        self.status_label.setText("Cancelando com segurança…")
        self.log("Cancelamento solicitado.")

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802 - API Qt
        if self.thread and self.thread.isRunning():
            if not self.operation_cancellable:
                QMessageBox.information(
                    self,
                    "Autorização em andamento",
                    "Conclua ou cancele a autorização no navegador. Esta janela poderá "
                    "ser fechada assim que o login terminar ou expirar.",
                )
                event.ignore()
                return
            answer = QMessageBox.question(
                self,
                "Operação em andamento",
                "Deseja cancelar a operação e fechar quando ela terminar com segurança?",
            )
            if answer == QMessageBox.Yes and self.worker:
                self.pending_close = True
                self.worker.request_cancel()
                self.status_label.setText("Cancelando antes de fechar…")
            event.ignore()
            return
        event.accept()


def main(*, smoke_test: bool = False) -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("DriveAutomate")
    window = DriveAutomateWindow()
    window.show()
    if smoke_test:
        QTimer.singleShot(500, app.quit)
    raise SystemExit(app.exec())


if __name__ == "__main__":
    main()
