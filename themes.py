"""
Shared UI theme for all MoveUp satellite windows.

Provides the MOVEUP_THEME color palette and a single apply_theme() function
that creates + applies a ttk theme. Previously, each satellite window
(mainExpiring, mainSamples, mainVelocity) had its own copy.

Usage::

    from themes import MOVEUP_THEME, apply_theme

    class MyWindow(tk.Toplevel):
        def __init__(self, master):
            super().__init__(master)
            apply_theme(self, "my_window_theme")
"""

from tkinter import ttk


# Kawaii-inspired lavender/purple palette shared across all satellite windows.
MOVEUP_THEME = {
    "bg": "#EEEAF8",
    "label_fg": "#3A2869",
    "tree_bg": "#F6F4FC",
    "tree_sel": "#C3B1E8",
    "btn_bg": "#EDEAF7",
    "btn_bg_active": "#DED9F0",
    "btn_fg": "#1F2328",
    "btn_border": "#7251A8",
}


def apply_theme(window, theme_name: str = "moveup_satellite") -> None:
    """
    Apply the shared ``MOVEUP_THEME`` color palette to a Tk/Toplevel window.

    Creates a named ttk theme using ``ttk.Style.theme_create()`` (if it does
    not already exist) then activates it with ``style.theme_use()``.  The
    theme is derived from the ``"clam"`` parent theme, which supports all the
    widget element overrides used here (background, foreground, map states).

    The ``clam`` parent is resolved by temporarily switching to it; if ``clam``
    is unavailable the currently active theme is used as the parent instead
    (the ``try/except`` handles platforms where ``clam`` may not be bundled).

    Styled widget classes: ``TFrame``, ``TLabel``, ``TCheckbutton``,
    ``TNotebook``, ``TNotebook.Tab``, ``TButton`` (with active/pressed map),
    ``Treeview``, ``Treeview.Heading``, ``TEntry``, ``TSpinbox``.

    Parameters
    ----------
    window : tk.Toplevel | tk.Tk
        The window whose ``background`` option is also set to ``MOVEUP_THEME["bg"]``
        so the bare Tk window frame matches the styled widgets.
    theme_name : str
        Unique ttk theme name.  Each satellite window should pass a distinct
        name (e.g. ``"expiring_theme"``, ``"velocity_theme"``) so that multiple
        satellite windows open simultaneously do not share/override each other's
        style instances.  ``theme_create`` is a no-op if the name already exists.
    """
    style = ttk.Style()
    try:
        base = style.theme_use()
        style.theme_use("clam")
    except Exception:
        base = style.theme_use()

    t = MOVEUP_THEME

    if theme_name not in style.theme_names():
        style.theme_create(
            theme_name,
            parent=base,
            settings={
                "TFrame": {"configure": {"background": t["bg"]}},
                "TLabel": {"configure": {"background": t["bg"], "foreground": t["label_fg"]}},
                "TCheckbutton": {"configure": {"background": t["bg"], "foreground": t["label_fg"]}},
                "TNotebook": {"configure": {"background": t["bg"]}},
                "TNotebook.Tab": {
                    "configure": {"background": t["bg"], "foreground": t["label_fg"], "padding": (12, 6)},
                },
                "TButton": {
                    "configure": {
                        "padding": 6, "relief": "solid", "borderwidth": 1,
                        "background": t["btn_bg"], "foreground": t["btn_fg"],
                    },
                    "map": {
                        "background": [("active", t["btn_bg_active"]), ("pressed", t["btn_bg_active"])],
                        "foreground": [("active", t["btn_fg"]), ("pressed", t["btn_fg"])],
                    },
                },
                "Treeview": {
                    "configure": {
                        "background": t["tree_bg"], "fieldbackground": t["tree_bg"],
                        "foreground": "#333333", "rowheight": 24,
                    },
                    "map": {
                        "background": [("selected", t["tree_sel"])],
                        "foreground": [("selected", "#000000")],
                    },
                },
                "Treeview.Heading": {
                    "configure": {"background": t["bg"], "foreground": t["label_fg"]},
                },
                "TEntry": {"configure": {"fieldbackground": "white", "foreground": "#1F2328"}},
                "TSpinbox": {"configure": {"fieldbackground": "white", "foreground": "#1F2328"}},
            },
        )

    style.theme_use(theme_name)
    window.configure(background=t["bg"])
