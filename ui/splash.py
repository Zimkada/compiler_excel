"""
Écran de démarrage (splash) premium, dessiné au thème — aucune image externe.
"""

from PyQt6.QtWidgets import QSplashScreen
from PyQt6.QtCore import Qt, QRect
from PyQt6.QtGui import QPixmap, QPainter, QColor, QFont, QBrush, QPen

from ui.styles import theme as T
from utils import resource_path


def make_splash() -> QSplashScreen:
    """Crée un splash screen 460×280 aux couleurs du design system."""
    w, h = 460, 280
    pix = QPixmap(w, h)
    pix.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pix)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    # Carte de fond arrondie
    painter.setBrush(QBrush(QColor(T.BG_SURFACE)))
    painter.setPen(QPen(QColor(T.BORDER), 1))
    painter.drawRoundedRect(1, 1, w - 2, h - 2, T.RADIUS_LG, T.RADIUS_LG)

    # Bandeau accent en haut
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QBrush(QColor(T.ACCENT)))
    painter.drawRoundedRect(0, 0, w, 96, T.RADIUS_LG, T.RADIUS_LG)
    painter.drawRect(0, 48, w, 48)  # bas du bandeau droit

    # Pastille logo : vrai logo de l'app (icon.ico) avec repli emoji
    painter.setBrush(QBrush(QColor(255, 255, 255, 45)))
    painter.drawRoundedRect(w // 2 - 28, 24, 56, 56, T.RADIUS_MD, T.RADIUS_MD)
    logo_path = resource_path("icon.ico")
    logo_pixmap = QPixmap(str(logo_path)) if logo_path.exists() else QPixmap()
    if not logo_pixmap.isNull():
        scaled = logo_pixmap.scaled(
            40, 40,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        target = QRect(w // 2 - 28, 24, 56, 56)
        x = target.x() + (target.width() - scaled.width()) // 2
        y = target.y() + (target.height() - scaled.height()) // 2
        painter.drawPixmap(x, y, scaled)
    else:
        painter.setPen(QColor(T.TEXT_ON_ACCENT))
        painter.setFont(QFont(T.FONT_FAMILY, 24))
        painter.drawText(w // 2 - 28, 24, 56, 56, Qt.AlignmentFlag.AlignCenter, "📊")

    # Titre
    painter.setPen(QColor(T.TEXT_PRIMARY))
    painter.setFont(QFont(T.FONT_FAMILY, 20, QFont.Weight.Bold))
    painter.drawText(0, 120, w, 36, Qt.AlignmentFlag.AlignCenter, "ExcelCompiler")

    # Accroche
    painter.setPen(QColor(T.TEXT_SECONDARY))
    painter.setFont(QFont(T.FONT_FAMILY, 10))
    painter.drawText(0, 158, w, 24, Qt.AlignmentFlag.AlignCenter,
                     "Compilateur Excel intelligent")

    # Pied : chargement + version
    painter.setPen(QColor(T.TEXT_MUTED))
    painter.setFont(QFont(T.FONT_FAMILY, 9))
    painter.drawText(0, h - 44, w, 20, Qt.AlignmentFlag.AlignCenter,
                     "Chargement…")
    painter.drawText(0, h - 26, w, 18, Qt.AlignmentFlag.AlignCenter, "v3.2")

    painter.end()

    splash = QSplashScreen(pix, Qt.WindowType.WindowStaysOnTopHint)
    splash.setMask(pix.mask())
    return splash
