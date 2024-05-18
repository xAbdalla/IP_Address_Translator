import sys
import gui


def main():
    # Show splash screen
    try:
        if getattr(sys, 'frozen', False):
            import pyi_splash  # type: ignore
    finally:
        pass

    app = gui.GUI(log=True)

    # Close splash screen
    try:
        if getattr(sys, 'frozen', False):
            pyi_splash.close()
    finally:
        pass

    app.mainloop()


if __name__ == '__main__':
    main()
