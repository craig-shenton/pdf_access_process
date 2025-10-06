"""Application entry point bootstrapping the Tkinter UI."""
from __future__ import annotations

from ui.app import PdfToAccessApp


def main() -> None:
    app = PdfToAccessApp()
    app.mainloop()


if __name__ == "__main__":
    main()
