"""
Bisa — Animated ASCII companion for the Move-Up Utility.

She's a cat companion who reacts to user interactions
and app events with various animations, tricks, and seasonal themes.
"""

import random
from datetime import datetime
import tkinter as tk
from tkinter import simpledialog


class AsciiDogWidget:
    """Animated ASCII companion widget — Bisa the cat.

    - click her to pet (receive_pet)
    - click box/blank space to throw her a treat (throw_treat_at_window_x / frame click)
    - stats counter (pets/treats)
    - react_* methods used by the app
    - idle micro-animations (wag, blink, sleep, zoomies)
    - contextual reactions (success/warning/error)
    - seasonal theme accents (Oct/Dec)
    - rare "legendary" easter egg (1% chance)
    """

    # ------------------------------
    # Existing frames
    # ------------------------------
    IDLE_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( O.O ) \n  > W <\n /|   |\\\n(_|   |_)",
    ]
    PET_FRAMES = [
        "  /\\_/\\  \n ( u.u )\u2665\n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^w^ )\u2665\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Nuzzle into hand
        "  /\\_/\\ \u2665\n ( ^.^ )~\n  > ^~<\n /|   |\\\n(_|   |_)",
        # Purr vibrate
        "  /\\_/\\  \n~( ^w^ )~\u2665\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Lean in
        "  /\\_/\\  \n ( ^.^ )\u2665\n  > ^~<\n /|   |\\\n(_|   |_)",
        # So content
        "  /\\_/\\  \n ( u.u )\u2665\n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( u.u ) \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    TREAT_SHORT = [
        "  /\\_/\\    \U0001f9b4\n ( o.o )  \n  > ^ <\n /|   |\\\n(_|   |_)",
        "    /\\_/\\ \U0001f9b4\n   ( ^.^) \n    > ^ <\n   /|   |\\\n  (_|   |_)",
    ]
    TREAT_MEDIUM = [
        "  /\\_/\\      \U0001f9b4\n ( o.o )    \n  > ^ <\n /|   |\\\n(_|   |_)",
        "    /\\_/\\  \U0001f9b4\n   ( ^o^ ) \n    > ^ <\n   /|   |\\\n  (_|   |_)",
        "      /\\_/\\\U0001f9b4\n     ( ^.^)\n      > ^ <\n     /|   |\\\n    (_|   |_)",
    ]
    TREAT_FAR = [
        "  /\\_/\\        \U0001f9b4\n ( o.o )      \n  > ^ <\n /|   |\\\n(_|   |_)",
        "    /\\_/\\    \U0001f9b4\n   ( ^o^ )   \n    > ^ <\n   /|   |\\\n  (_|   |_)",
        "      /\\_/\\ \U0001f9b4\n     ( ^.^) \n      > ^ <\n     /|   |\\\n    (_|   |_)",
        "        /\\_/\\\U0001f9b4\n       ( ^O^)\n        > ^ <\n       /|   |\\\n      (_|   |_)",
    ]
    RUN_BACK = [
        "      /\\_/\\  \n     \U0001f9b4(^.^) \n      > ^ <\n     /|   |\\\n    (_|   |_)",
        "    /\\_/\\    \n   \U0001f9b4( ^w^)  \n    > ^ <\n   /|   |\\\n  (_|   |_)",
        "  /\\_/\\      \n \U0001f9b4( ^.^)   \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    HAPPY_FRAMES = [
        "    /\\_/\\   \n   ( ^o^)o  \n    > ^ <\n   /|   |\\\n  (_|   |_)",
        "      /\\_/\\ \n     ( ^o^)o\n      > ^ <\n     /|   |\\\n    (_|   |_)",
        "    /\\_/\\   \n   o(^w^)   \n    > ^ <\n   /|   |\\\n  (_|   |_)",
        # Jump up!
        "    /\\_/\\   \n   \\(^o^)/  \n    > ^ <\n      \n   ~~~~~~~~",
        "    /\\_/\\   \n   ( ^.^)o  \n    >w^ <\n   /|   |\\\n  (_|   |_)",
        "      /\\_/\\ \n     o(^v^) \n      > ^ <\n     /|   |\\\n    (_|   |_)",
        # Twirl!
        "    /\\_/\\   \n   \\(^o^)/o \n    > ^ <\n   /|   |\\\n  (_|   |_)",
        # Happy wiggle
        "    /\\_/\\  ~\n   o( ^w^ ) \n    > ^ <\n   /|   |\\\n  (_|   |_)",
    ]
    LOAD_FRAMES = [
        "  /\\_/\\   \n ( O.O )! \n  > W <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\   \n ( ^o^)!! \n  >w^ <\n /|   |\\\n(_|   |_)",
        "   /\\_/\\  \n  (\\^o^/) \n   >W< \n  /|   |\\\n (_|   |_)",
        "  /\\_/\\   \n  (^o^)/  \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\   \n \\(^w^)/  \n  >w^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\   \n  ( ^.^)~ \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    EXCLUDED_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( o.o )!\n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ;.; ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( T.T ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        # Looks down sadly
        "  /\\_/\\  \n ( T.T ) \n  > v <\n /|   |\\\n(_|   |_)",
        # Accepts it
        "  /\\_/\\  \n ( u.u ) \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    ALERT_FRAMES = [
        "  /|_|\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /|_|\\  \n ( O.O ) \n  > ! <\n /|   |\\\n(_|   |_)",
        "  /|_|\\  \n ( ^.^ ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    SNIFF_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( o.~ ) \n  >sniff\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^.o ) \n  >sniff\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    KUNTAL_FRAMES = [
        "  /\\_/\\  \n ( ^o^)\u2605\n  > ^ <\n /|   |\\\n(_|   |_)",
        "   /\\_/\\ \n  (\u2605^o^)\n   > ^ <\n  /|   |\\\n (_|   |_)",
        "  /\\_/\\  \n  (^w^)\u2605\n  >w^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n \\(^o^)/ \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n  ( ^.^)\u2605\n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    STRETCH_FRAMES = [
        # Waking up groggy
        "  /\\_/\\  \n ( -.- ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        # Big yawn!
        "  /\\_/\\  \n ( o.O ) \n  > o <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( O.O ) \n  > O <\n /|   |\\\n(_|   |_)",
        # Front stretch (butt up)
        "  /\\_/\\  \n ( o.o ) \n  >str<\n /|   |\\\n(_|   |_)",
        # Back arch
        "  /\\_/~  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        # Toe beans spread
        "  /\\_/\\  \n ( ^.^ ) \n  >\\|/<\n /|   |\\\n(_|   |_)",
        # Done! Feels good
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    CLEARED_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( -.- ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( u.u ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]

    # ------------------------------
    # New frames (added)
    # ------------------------------
    WAG_FRAMES = [
        "  /\\_/\\  ~\n ( ^.^ )  \n  > ^ <\n /|   |\\\n(_|   |_)",
        "~  /\\_/\\  \n  ( ^.^ ) \n   > ^ <\n  /|   |\\\n (_|   |_)",
        "  /\\_/\\  ~\n ( ^w^ )  \n  > ^ <\n /|   |\\\n(_|   |_)",
        "~  /\\_/\\  \n  ( ^w^ ) \n   > ^ <\n  /|   |\\\n (_|   |_)",
    ]
    BLINK_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( -.- ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    SLEEP_FRAMES = [
        "  /\\_/\\  \n ( -.- ) zZ\n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( -.- ) Zz\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Deeper sleep
        "  /\\_/\\  \n ( -.- ) zz\n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( u.u ) zZ\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Dream bubble!
        "  /\\_/\\ \U0001f4ad\n ( -w- ) Zz\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Ear twitch in sleep
        "  /|_/\\  \n ( -.- ) zZ\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Peaceful
        "  /\\_/\\  \n ( -w- ) Zz\n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    ZOOMIES_FRAMES = [
        "  /\\_/\\      \n ( ^o^ )  \u26a1\n  > ^ <      \n /|   |\\\n(_|   |_)",
        "      /\\_/\\  \n \u26a1 ( ^o^ ) \n      > ^ <  \n     /|   |\\\n    (_|   |_)",
        # Wall bounce!
        "        /\\_/\\\n       ( >o< )|\n        > ^ < |\n       /|   |\\\n      (_|   |_)",
        "  /\\_/\\      \n ( ^w^ )  \u26a1\n  > ^ <      \n /|   |\\\n(_|   |_)",
        "    /\\_/\\    \n \u26a1 ( ^w^ ) \n    > ^ <    \n   /|   |\\\n  (_|   |_)",
        # Bounce other way!
        "|/\\_/\\      \n|( >o< ) \u26a1  \n| > ^ <      \n /|   |\\\n(_|   |_)",
        # Skid to stop
        "  /\\_/\\  ~~\n ( ^o^ )  \u26a1\n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    CONFUSED_FRAMES = [
        "  /\\_/\\  \n ( o.o ) ?\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Head tilt right
        "  /\\_/\\  \n ( O.o ) ?\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Head tilt left
        "  /\\_/\\  \n ( o.O ) ?\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Double question
        "  /\\_/\\  \n ( o.o ) ??\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Gives up trying to understand
        "  /\\_/\\  \n ( -.- ) ~\n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    BELLY_FRAMES = [
        "  /\\_/\\  \n  ( ^o^ ) \n  ~> w <~\n  \\|   |/\n  (_|_|_)",
        "  /\\_/\\  \n  ( ^w^ )\u2665\n  ~> w <~\n  \\|   |/\n  (_|_|_)",
        # Air-kneading (making biscuits!)
        "  /\\_/\\  \n  ( ^w^ )\u2665\n  ~>\\w/<~\n  \\|   |/\n  (_|_|_)",
        # Purr kick
        "  /\\_/\\  \n  ( ^o^ )\u2665\n  ~> w <~\n  \\| ~/|/\n  (_|_|_)",
        "  /\\_/\\  \n  ( ^.^ )\u2665\n  ~> w <~\n  \\|   |/\n  (_|_|_)",
        # So happy, eyes closing
        "  /\\_/\\  \n  ( u.u )\u2665\n  ~> ^ <~\n  \\|   |/\n  (_|_|_)",
        # Purring bliss
        "  /\\_/\\  \n ~( -w- )\u2665\n  ~> ^ <~\n  \\|   |/\n  (_|_|_)",
    ]
    SUCCESS_FRAMES = [
        "    /\\_/\\   \n   ( ^o^)\u2728 \n    > ^ <\n   /|   |\\\n  (_|   |_)",
        "      /\\_/\\ \n     ( ^w^)\u2728\n      > ^ <\n     /|   |\\\n    (_|   |_)",
        # Fist pump!
        "    /\\_/\\   \n   \\(^o^)/\u2728\n    > ^ <\n   /|   |\\\n  (_|   |_)",
        # Sparkle shimmy
        "  \u2728/\\_/\\\u2728 \n   ( ^w^ )\u2728\n    > ^ <\n   /|   |\\\n  (_|   |_)",
        # Proud stance
        "    /\\_/\\   \n   ( ^.^ )\u2728\n    > ^ <\n   /|   |\\\n  (_|   |_)",
    ]
    WARNING_FRAMES = [
        # Ears perk up
        "  /|_|\\  \n ( O.O ) !\n  > ! <\n /|   |\\\n(_|   |_)",
        "  /|_|\\  \n ( O.O )!!\n  > ! <\n /|   |\\\n(_|   |_)",
        "  /|_|\\  \n ( o.o ) !\n  > ! <\n /|   |\\\n(_|   |_)",
        # Backing away
        "  /\\_/\\  \n ( o.o ) !\n  > ! <\n /|   |\\\n(_|   |_)",
        # Cautious
        "  /\\_/\\  \n ( o.o )  \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    LEGENDARY_FRAMES = [
        # Power building...
        "  /\\_/\\    \u2605\n ( \u2727o\u2727 )  \u2605\n  > W <   \n /|   |\\\n(_|   |_)",
        # More stars!
        "  /\\_/\\   \u2605\u2605\n ( \u2727o\u2727 )  \u2605\n  > W <   \u2605\n /|   |\\\n(_|   |_)",
        # FULL GLORY
        "  /\\_/\\   \u2605\u2605\u2605\n ( \u2727w\u2727 )  \u2605\n  > W <   \u2605\n /|   |\\\n(_|   |_)",
        "\u2605 /\\_/\\  \u2605\u2605\u2605\n\u2605( \u2727o\u2727 ) \u2605\u2605\n  > W <   \u2605\n /|   |\\\n(_|   |_)",
        # Sparkling pose
        "  /\\_/\\   \u2605\u2605\u2605\n \\(\u2727.^\u2727)/  \u2605\n  > W <   \u2605\n /|   |\\\n(_|   |_)",
        # Glowing
        "\u2605 /\\_/\\  \u2605\u2605\u2605\n\u2605(\u2727w\u2727)  \u2605\u2605\n  > W < \u2605 \u2605\n /|   |\\\n(_|   |_)",
        # Legendary stance
        "  /\\_/\\   \u2605\u2605\u2605\n ( \u2727o\u2727 )  \u2605\n  > W <   \u2605\n /|   |\\\n(_|   |_)",
    ]

    # Trick frames
    SIT_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^.^ ) \n  |   |\n  |   | \n  |___| ",
        "  /\\_/\\  \n ( ^.^ ) \n  | W |\n  |   | \n  |___| ",
        "  /\\_/\\  \n ( u.u ) \n  | W |\n  |   | \n  |___| ",
    ]
    SHAKE_FRAMES = [
        "  /\\_/\\  \n ( o.o )/ \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^.^ )\U0001f91d\n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^w^ )\U0001f91d\n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\  \n ( ^.^ )/ \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    SPIN_FRAMES = [
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "   (\\  \n    )  \n   /  \n   |  \n  (_) ",
        "  \\_/\\_ \n   ( \u25cf ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        "      /)\n     (  \n      \\ \n      |  \n     (_) ",
        "  /\\_/\\  \n ( ^o^ )~\n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    PLAY_DEAD_FRAMES = [
        # Dramatic gasp!
        "  /\\_/\\  \n ( O.O )!\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Stagger...
        "   /\\_/\\ \n  ( x.x )\n   > ^ <\n  /|   |\\\n (_|   |_)",
        # Fall over
        "  /\\_/\\  \n ( x.x ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        # Fully dead
        "             \n /\\_/\\_____\n ( x.x      )\n  |__|  |__|\n ~~~~~~~~~~~",
        # Leg twitch
        "             \n /\\_/\\_____\n ( x.x  ~   )\n  |__|  |__|\n ~~~~~~~~~~~",
        # ...peeking?
        "             \n /\\_/\\_____\n ( -.x  ~   )\n  |__|  |__|\n ~~~~~~~~~~~",
    ]
    SNEEZE_FRAMES = [
        # Nose tickle
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        # Building up...
        "  /\\_/\\  \n ( o.O ) \n  > o <\n /|   |\\\n(_|   |_)",
        # Ahhh...
        "  /\\_/\\  \n ( O.O ) \n  > O <\n /|   |\\\n(_|   |_)",
        # ACHOOOO!!
        "  /\\_/\\  \n (>w< )!! \n  > ^ < !!\n /|   |\\\n(_|   |_)",
        # Recoil
        "   /\\_/\\ \n  ( >.< )~\n   > ^ <\n  /|   |\\\n (_|   |_)",
        # Dazed recovery
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ < ~\n /|   |\\\n(_|   |_)",
    ]

    HALLOWEEN_FRAMES = [
        "  /\\_/\\   \U0001f383\n ( o.o )  \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\   \U0001f383\n ( O.O )  \n  > W <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\   \U0001f47b\n ( ^.^ )  \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]
    WINTER_FRAMES = [
        "  /\\_/\\   \u2744\ufe0f\n ( o.o )  \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\   \u2744\ufe0f\n ( ^.^ )  \n  > ^ <\n /|   |\\\n(_|   |_)",
        "  /\\_/\\   \u2603\ufe0f\n ( u.u )  \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]

    BOX_FRAMES = [
        # Bisa spots the box
        "  /\\_/\\  \n ( O.O )  \U0001f4e6\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Investigating...
        "   /\\_/\\ \U0001f4e6\n  ( o.o )\n   > ^ <\n  /|   |\\\n (_|   |_)",
        # Sniff sniff...
        "   /\\_/\\\U0001f4e6\n  ( o.~ )\n   >sniff\n  /|   |\\\n (_|   |_)",
        # One paw in...
        "  _____\n  /\\_/\\|\n ( o.o )|\n  > ^ < |\n  |_____|",
        # Climbing in...
        "  _____\n /\\_/\\ |\n ( ^.^ )|\n  > ^ < |\n  |_____|",
        # Squeezing in!
        "  _____\n | /\\  |\n |(^.^)|\n | > < |\n  |_____|",
        # If I fits...
        "  _____\n |/\\_/\\|\n |( ^.^)|\n |     |\n  |_____|",
        # Wiggle wiggle~
        "  _____\n |/\\_/\\|\n |( ^w^)|\n |  ~  |\n  |_____|",
        # I sits!!
        "  _____\n |/\\_/\\|\n |( u.u)|\n |     |\n  |_____|",
        # So comfy... zzz
        "  _____\n |/\\_/\\|\n |( -.-)|zZ\n |     |\n  |_____|",
    ]

    BOWLING_FRAMES = [
        # Picks up the ball
        "  /\\_/\\  \n ( ^.^ )\U0001f3b3\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Lines up — pins ahead!
        "  /\\_/\\        i\n ( -.- )\U0001f3b3   i i\n  > ^ <      i i i\n /|   |\\\n(_|   |_)",
        # Winds up...
        "  /\\_/\\        i\n\U0001f3b3( >.<)    i i\n  > ^ <      i i i\n /|   |\\\n(_|   |_)",
        # ROLLS IT!
        "  /\\_/\\        i\n ( >o< )/   i i\n  > ^ <      i i i\n /|   |\\\n(_|   |_)",
        # Ball rolling toward pins...
        "  /\\_/\\    \U0001f3b3  i\n ( o.o )     i i\n  > ^ <      i i i\n /|   |\\\n(_|   |_)",
        # Almost there!
        "  /\\_/\\      \U0001f3b3i\n ( O.O )     i i\n  > ^ <      i i i\n /|   |\\\n(_|   |_)",
        # CRASH!! Pins scatter!
        "  /\\_/\\    \U0001f4a5  \n ( ^o^ )   i \\ /i\n  > ^ <     /i\\ \n /|   |\\\n(_|   |_)",
        # STRIKE!!!
        "  /\\_/\\  \n ( ^o^ )  STRIKE!\n  > ^ <   \u2728\U0001f3b3\u2728\n /|   |\\\n(_|   |_)",
        # Victory dance!
        "  /\\_/\\  \n \\(^w^)/  \u2728\U0001f3b3\u2728\n  > ^ <\n /|   |\\\n(_|   |_)",
    ]

    LASER_FRAMES = [
        # Spots the dot...
        "  /\\_/\\  \n ( O.O )  \U0001f534\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Eyes lock on — butt wiggle
        "  /\\_/\\   \U0001f534\n ( O.O )  \n  > ^ <\n  |   | \n  |___|  ",
        # Dot moves left!
        "\U0001f534 /\\_/\\  \n  ( O.O ) \n   > ^ <\n  /|   |\\\n (_|   |_)",
        # POUNCE left!
        "\U0001f534/\\_/\\  \n ( >o< ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        # Dot escapes right!
        "  /\\_/\\  \n ( O.O )  \n  > ^ <  \U0001f534\n /|   |\\\n(_|   |_)",
        # Chase right!
        "      /\\_/\\\U0001f534\n     ( >.<) \n      > ^ <\n     /|   |\\\n    (_|   |_)",
        # Dot on head?!
        "  \U0001f534\\_/\\  \n ( O.O ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        # POUNCE on it!!
        "  /\\_/\\  \n ( >w< )\U0001f534\n  > ^ <\n  |   | \n  |___|  ",
        # ...where'd it go?
        "  /\\_/\\  \n ( o.o ) ?\n  > ^ <\n /|   |\\\n(_|   |_)",
    ]

    DAISY_FRAMES = [
        # Sniffs... what's that?
        "  /\\_/\\  \n ( o.o ) \n  > ^ <\n /|   |\\\n(_|   |_)",
        # A daisy appears!
        "  /\\_/\\ \U0001f33c\n ( ^.^ )  \n  > ^ <\n /|   |\\\n(_|   |_)",
        # More daisies blooming!
        "\U0001f33c/\\_/\\ \U0001f33c\n ( ^o^ )  \n  > ^ <\n /|   |\\\n(_|   |_)",
        # Flowers everywhere!
        "\U0001f33c/\\_/\\ \U0001f33c\n ( ^o^ ) \U0001f338\n \U0001f33c> ^ <\n /|   |\\\n(_|   |_)",
        # Happy dance with flowers
        "\U0001f33c /\\_/\\\U0001f33c\n  \\(^w^)/\U0001f338\n \U0001f33c > ^ <\U0001f33c\n  /|   |\\\n (_|   |_)",
        # Twirling in daisies!
        "\U0001f338/\\_/\\ \U0001f33c\n\U0001f33c\\(^o^)/\U0001f338\n \U0001f33c> ^ <\U0001f33c\n /|   |\\\n(_|   |_)",
        # Full bloom garden!
        "\U0001f33c\U0001f338/\\_/\\\U0001f33c\U0001f338\n\U0001f33c( ^w^ )\U0001f33c\n\U0001f338 >w^ <\U0001f338\n /|   |\\\n(_|   |_)",
        # Dancing in flowers~
        "\U0001f338\U0001f33c/\\_/\\\U0001f338\U0001f33c\n\U0001f33c\\(^o^)/\U0001f338\n\U0001f33c >w^ <\U0001f33c\n /|   |\\\n(_|   |_)",
        # Flower shower!!
        "\U0001f33c\U0001f338\U0001f33c\U0001f338\U0001f33c\n\U0001f338/\\_/\\\U0001f33c\n ( ^w^ )\U0001f338\n\U0001f33c>w^ <\U0001f33c\n /|   |\\\U0001f338",
        # Happy in the garden~
        "\U0001f33c\U0001f338/\\_/\\\U0001f33c\U0001f338\n\U0001f33c( u.u )\U0001f33c\n\U0001f338 > ^ <\U0001f338\n /|   |\\\n(_|   |_)",
    ]

    # Phase 1: Discovery & consumption (fast, excited)
    CATNIP_DISCOVER = [
        # Spots the nug
        "  /\\_/\\  \n ( O.O ) \U0001f33f\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Sniff sniff...
        "  /\\_/\\  \U0001f33f\n ( o.~ )  \n  >sniff\n /|   |\\\n(_|   |_)",
        # oh YES
        "  /\\_/\\ \U0001f33f\n ( \u00d8.\u00d8 )!\n  > o <\n /|   |\\\n(_|   |_)",
        # NOM
        "  /\\_/\\  \n ( >w< )\U0001f33f\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Chewing...
        "  /\\_/\\  \n ( ^w^ ) \n  > ^ <\n /|   |\\\n(_|   |_)",
    ]

    # Phase 2: The high kicks in (medium, trippy chaos)
    CATNIP_HIGH = [
        # Pupils dilate
        "  /\\_/\\ \U0001f4a8\n ( \u00d8.\u00d8 )\n  > o <\n /|   |\\\n(_|   |_)",
        # Whoa...
        " ~/\\_/\\~\U0001f4a8\n~( *.* )~\n  > W <\n /|   |\\\n(_|   |_)",
        # Rolling
        "    /\\_/\\\U0001f4a8\n   ( @.@ ) \n    > ^ <\n   /|   |\\\n  (_|   |_)",
        # Seeing things
        "\U0001f4a8/\\_/\\  \n ( *.* ) \u2728\n ~> W <~\n /|   |\\\n(_|   |_)",
        # Wiggly
        " ~/\\_/\\~\U0001f4a8\n~( @.@ )~\n ~> W <~\n /|   |\\\n(_|   |_)",
        # Tumble
        "/\\_/\\  \U0001f4a8\n( *.* )  \n > ^ <\n/|   |\\\n(_|   |_)",
        # Vibing hard
        "  \U0001f4a8/\\_/\\\n    ( @w@ )\n     > o <\n    /|   |\\\n   (_|   |_)",
        # Peak eyes
        " ~\U0001f4a8\\_/\\~ \n~( \u00d8.\u00d8 )~\U0001f4a8\n ~> W <~\n /|   |\\\n(_|   |_)",
    ]

    # Phase 3: Stoner couch lock (slow, chill)
    CATNIP_CHILL = [
        # Sinking in...
        "\U0001f4a8 /\\_/\\ \U0001f4a8\n  ( -.- ) \n   > ^ <\n  /|   |\\\n (_|   |_)",
        # Couch lock
        "\U0001f33f /\\_/\\ \U0001f33f\n  ( -w- ) \n   > ^ <\n   |   | \n   |___| ",
        # duuude...
        "\U0001f4a8 /\\_/\\ \U0001f33f\n  ( \u00b0.\u00b0 ) \n   > ^ <\n   |   | \n   |___| ",
        # So chill
        "\U0001f33f /\\_/\\ \U0001f4a8\n  ( \u00b0w\u00b0 )\n   > ^ <\n   |   | \n   |___| ",
        # Munchies!!
        "\U0001f33f /\\_/\\ \U0001f355\n  ( ^o^ )nom\n   > ^ <\n   |   | \n   |___| ",
        # More munchies
        "\U0001f4a8 /\\_/\\ \U0001f36a\n  ( ^w^ )nom\n   > ^ <\n   |   | \n   |___| ",
        # Spacing out
        "\U0001f33f /\\_/\\ \U0001f4a8\n  ( \u00b0.\u00b0 ) \n   > ^ <\n   |   | \n   |___| ",
        # So good...
        "\U0001f4a8 /\\_/\\ \U0001f33f\n  ( -w- )\u2665\n   > ^ <\n   |   | \n   |___| ",
    ]

    # Phase 4: Recovery wobble (before zoomies)
    CATNIP_RECOVER = [
        # Waking up...
        "  /\\_/\\ \U0001f4a8\n ( @.@ ) \n  > ^ <\n  \\|   |/\n  (_|_|_)",
        # Blink blink
        "  /\\_/\\  \n ( o.o )~\U0001f4a8\n  > ^ <\n /|   |\\\n(_|   |_)",
        # Standing up wobbly
        "  /\\_/\\  \n ( ^.^ )~\n  > ^ <\n /|   |\\\n(_|   |_)",
    ]

    MESSAGES = {
        "idle":     "...",
        "pet":      "so nice~ \u2665",
        "treat":    "treat?? \U0001f9b4",
        "running":  "nom nom! \U0001f9b4",
        "happy":    "yay!!!! \u2728",
        "loaded":   "new data!! \U0001f4cb",
        "excluded": "oh no... \U0001f622",
        "sniff":    "sniff sniff...",
        "alert":    "! what's that?",
        "kuntal":   "ooh priority! \u2605",
        "stretch":  "zzz... yawn~",
        "cleared":  "phew~ clean!",
        "restored": "yay, back!! \u2705",
        "wag":      "purrrr~!!",
        "blink":    "blink~",
        "sleep":    "zzz\u2026",
        "zoomies":  "ZOOMIES!! \u26a1",
        "confused": "huh?",
        "success":  "nice!! \u2705",
        "warning":  "uh oh\u2026 \u26a0\ufe0f",
        "error":    "nope\u2026 \U0001f4a5",
        "legendary": "LEGENDARY BISAAAA \u2605\u2605\u2605",
        "halloween": "spooky Bisa \U0001f383",
        "winter":   "brr\u2026 \u2744\ufe0f",
        "belly":    "belly rubs!! \u2665",
        "milestone": "milestone!!  \u2b50",
        "moveup":    "they moved!! \U0001f4e6",
        "sit":       "good sit!! \U0001f43e",
        "shake":     "nice to meet u! \U0001f91d",
        "spin":      "wheee~! \U0001f300",
        "play_dead": "... \U0001f480 (jk!!)",
        "sneeze":    "ACHOO!! \U0001f927",
        "box":       "if I fits... \U0001f4e6",
        "bowling":   "STRIKE!! \U0001f3b3",
        "laser":     "RED DOT!! \U0001f534",
        "daisy":     "flowers!! \U0001f33c\U0001f338",
        "catnip":    "CATNIP!! \U0001f33f\U0001f4a8",
        "catnip_high": "whooaaa... \U0001f4a8\U0001f33f\U0001f4a8",
        "catnip_chill": "duuude... \U0001f33f\U0001f4a8",
        "catnip_munch": "munchies!! \U0001f355\U0001f36a",
        "catnip_earn": "earned catnip! \U0001f33f",
    }

    THOUGHTS = [
        "thinking about treats...",
        "I wonder what's in Backstock...",
        "is it lunch yet? \U0001f355",
        "so many barcodes...",
        "~dreaming of zoomies~",
        "*stares at spreadsheet*",
        "who's a good kitty? me??",
        "need... more... pets...",
        "what does METRC even mean",
        "purr loading... 10%",
        "\u2728 sparkle sparkle \u2728",
        "hmm... sus barcode \U0001f50d",
        "inventory is my passion",
        "*pretends to help*",
        "one more export plz \U0001f4cb",
        "mrow? mrow.",
        "cannabis... the good stuff \U0001f33f",
        "*knocks barcode off desk*",
        "if I fits, I sits \U0001f4e6",
        "~napping in a box~",
        "is that... catnip? \U0001f33f",
    ]

    def __init__(self, parent: tk.Widget, name: str = "Bisa", on_rename=None):
        self.parent = parent
        self._name = name
        self._on_rename = on_rename
        self._state = "idle"
        self._after_id = None
        self._idle_idx = 0
        self._anim_idx = 0
        self._anim_frames = []
        self._total_pets = 0
        self._total_treats = 0
        self._total_moveups = 0
        self._interactions_since_milestone = 0
        self._next_milestone_interval = random.randint(60, 100)

        # Catnip reward system
        self._catnip_redeemed = 0     # persisted — how many have been used
        self._on_catnip_change = None  # callback for main.py to persist

        # Animation tuning
        self._speed_scale = 1.0
        self._legendary_chance = 0.01

        # Theme
        self._apply_seasonal_theme(datetime.now())

        self.frame = tk.Frame(
            parent,
            relief="ridge",
            bd=2,
            bg=self._theme_bg,
            highlightbackground=self._theme_border,
            highlightthickness=1,
            padx=8,
            pady=6,
        )

        self._title_label = tk.Label(
            self.frame,
            text=f"\u2726 {self._name} \u2726",
            font=("Segoe UI", 10, "bold"),
            bg=self._theme_bg,
            foreground=self._theme_accent,
            cursor="hand2",
        )
        self._title_label.pack()
        self._title_label.bind("<Double-Button-1>", lambda _e: self._show_rename_dialog())

        self.dog_var = tk.StringVar()
        self.dog_label = tk.Label(
            self.frame,
            textvariable=self.dog_var,
            font=("Courier", 11, "bold"),
            justify="center",
            cursor="hand2",
            bg=self._theme_bg,
            fg=self._theme_accent,
        )
        self.dog_label.pack(pady=(2, 0), fill="x", expand=True)
        self.dog_label.bind("<Button-1>", lambda _e: self.receive_pet())
        self.dog_label.bind("<Double-Button-1>", lambda _e: self._sneeze())
        self.dog_label.bind("<Button-3>", lambda _e: self._belly_rub())
        self.dog_label.bind("<Enter>", self._on_hover)

        # Secret trick input buffer
        self._trick_buffer = ""
        self.frame.bind("<Key>", self._on_key)
        self.frame.configure(takefocus=True)
        # Also let clicking the frame give it focus for key events
        self.frame.bind("<Button-1>", lambda e: (self.frame.focus_set(), self._on_frame_click(e)))

        self.msg_var = tk.StringVar(value="...")
        tk.Label(
            self.frame,
            textvariable=self.msg_var,
            font=("Segoe UI", 9),
            bg=self._theme_bg,
            fg=self._theme_msg,
        ).pack()

        tk.Frame(self.frame, bg=self._theme_border, height=1).pack(fill="x", pady=4)

        tk.Label(
            self.frame,
            text="click \u2192 pet  |  dbl-click \u2192 boop  |  right-click \u2192 belly rub  |  dbl-click name \u2192 rename",
            font=("Segoe UI", 8),
            bg=self._theme_bg,
            fg=self._theme_hint,
        ).pack(pady=(0, 2))

        self.stats_var = tk.StringVar(value="pets:0  treats:0")
        tk.Label(
            self.frame,
            textvariable=self.stats_var,
            font=("Segoe UI", 8),
            bg=self._theme_bg,
            fg=self._theme_stats,
        ).pack()

        self.catnip_var = tk.StringVar(value="")
        self._catnip_label = tk.Label(
            self.frame,
            textvariable=self.catnip_var,
            font=("Segoe UI", 9, "bold"),
            bg=self._theme_bg,
            fg="#2d8a4e",
            cursor="hand2",
        )
        self._catnip_label.pack(pady=(2, 0))
        self._catnip_label.bind("<Button-1>", lambda _e: self._redeem_catnip())

        self._random_trick_label = tk.Label(
            self.frame,
            text="\u2728 click for a trick! \u2728",
            font=("Segoe UI", 8),
            bg=self._theme_bg,
            fg=self._theme_accent,
            cursor="hand2",
        )
        self._random_trick_label.pack(pady=(2, 0))
        self._random_trick_label.bind("<Button-1>", lambda _e: self._play_random_trick())

        self._render_frame(self.IDLE_FRAMES[0])
        self._idle_loop()

    # ------------------------------
    # Theme
    # ------------------------------
    def _apply_seasonal_theme(self, now):
        # Defaults: your purple
        self._theme_bg = "#f0eaf4"
        self._theme_border = "#c9a8d4"
        self._theme_accent = "#7a4a9a"
        self._theme_msg = "#9c6dbf"
        self._theme_hint = "#c9a8d4"
        self._theme_stats = "#c9a8d4"

        self._seasonal_idle_frames = None

        # October: Halloween
        if now.month == 10:
            self._theme_bg = "#1f1326"
            self._theme_border = "#7a4a9a"
            self._theme_accent = "#ff7a18"
            self._theme_msg = "#c9a8d4"
            self._theme_hint = "#7f6a86"
            self._theme_stats = "#6a4b73"
            self._seasonal_idle_frames = self.HALLOWEEN_FRAMES

        # December: Winter
        elif now.month == 12:
            self._theme_bg = "#eef6ff"
            self._theme_border = "#b8d7ff"
            self._theme_accent = "#2a5aa5"
            self._theme_msg = "#3b76c9"
            self._theme_hint = "#8aa9d6"
            self._theme_stats = "#c7d9f2"
            self._seasonal_idle_frames = self.WINTER_FRAMES

    # ------------------------------
    # Core rendering
    # ------------------------------
    def _render_frame(self, text: str, msg: str = ""):
        self.dog_var.set(text)
        if msg:
            self.msg_var.set(msg)

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.parent.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    TITLES = [
        (0,    "New Kitten \U0001f431"),
        (10,   "Good Girl \U0001f408"),
        (50,   "Loyal Friend \U0001f43e"),
        (100,  "Treat Fiend \U0001f9b4"),
        (200,  "Inventory Cat \U0001f4e6"),
        (500,  "Store Guardian \U0001f6e1\ufe0f"),
        (1000, "LEGENDARY BISA \u2605"),
    ]

    def _get_title(self) -> str:
        total = self._total_pets + self._total_treats
        title = self.TITLES[0][1]
        for threshold, t in self.TITLES:
            if total >= threshold:
                title = t
        # Legendary title uses the pet's actual name
        if total >= 1000:
            title = f"LEGENDARY {self._name.upper()} \u2605"
        return title

    def _update_stats(self):
        title = self._get_title()
        self.stats_var.set(
            f"{title}  |  pets:{self._total_pets}  treats:{self._total_treats}  moved:{self._total_moveups}"
        )
        self._check_catnip_earned()

    def set_name(self, name: str):
        """Update the pet's display name."""
        self._name = name.strip() or "Bisa"
        self._title_label.config(text=f"\u2726 {self._name} \u2726")
        self._update_stats()

    def _show_rename_dialog(self):
        """Open a dialog to rename the pet. Triggered by double-clicking her name."""
        new_name = simpledialog.askstring(
            "Rename",
            f"Enter a new name for {self._name}:",
            initialvalue=self._name,
            parent=self.frame,
        )
        if new_name and new_name.strip():
            self.set_name(new_name.strip())
            self.msg_var.set(f"I'm {self._name} now!! \u2728")
            # Notify main app to persist the change
            if self._on_rename:
                self._on_rename(self._name)

    # ------------------------------
    # Catnip reward system
    # ------------------------------
    def _catnip_earned_total(self) -> int:
        """Calculate total catnip earned from lifetime stats."""
        from_interactions = (self._total_pets + self._total_treats) // 20
        from_moveups = self._total_moveups // 10
        return from_interactions + from_moveups

    def _catnip_available(self) -> int:
        """How many catnip treats are available to redeem."""
        return max(0, self._catnip_earned_total() - self._catnip_redeemed)

    def _update_catnip_display(self):
        """Update the catnip label to show available treats."""
        avail = self._catnip_available()
        if avail > 0:
            self.catnip_var.set(f"\U0001f33f catnip x{avail} — click to use!")
        else:
            self.catnip_var.set("")

    def _check_catnip_earned(self):
        """Check if a new catnip was just earned and flash a message."""
        old_avail = getattr(self, "_last_catnip_avail", 0)
        avail = self._catnip_available()
        self._last_catnip_avail = avail
        self._update_catnip_display()
        # Flash the earn message only when a NEW catnip appears
        if avail > old_avail and self._state == "idle":
            self.msg_var.set(self.MESSAGES["catnip_earn"])

    def _redeem_catnip(self):
        """Use one catnip treat — triggers full stoner cat experience."""
        if self._catnip_available() <= 0:
            return
        if self._state not in ("idle", "pet", "happy"):
            return
        self._catnip_redeemed += 1
        self._update_catnip_display()
        self._cancel()
        self._state = "catnip"
        # 5-phase catnip trip:
        #   1. Discovery & nom (fast, excited)
        #   2. The high kicks in (medium, trippy chaos)
        #   3. Stoner couch lock + munchies (slow, chill)
        #   4. Recovery wobble (medium)
        #   5. Post-catnip zoomies burst (fast)
        zoomies_burst = self.ZOOMIES_FRAMES * 3
        self._run_anim(
            self.CATNIP_DISCOVER,
            self.MESSAGES["catnip"],
            int(150 * self._speed_scale),
            lambda: self._run_anim(
                self.CATNIP_HIGH,
                self.MESSAGES["catnip_high"],
                int(220 * self._speed_scale),
                lambda: self._run_anim(
                    self.CATNIP_CHILL,
                    self.MESSAGES["catnip_chill"],
                    int(420 * self._speed_scale),
                    lambda: self._run_anim(
                        self.CATNIP_RECOVER,
                        "wh... what happened? \U0001f4a8",
                        int(320 * self._speed_scale),
                        lambda: self._run_anim(
                            zoomies_burst,
                            "ZOOMIES!! \u26a1\U0001f33f",
                            int(160 * self._speed_scale),
                            lambda: self._return_idle(),
                        ),
                    ),
                ),
            ),
        )
        # Notify main app to persist
        if self._on_catnip_change:
            self._on_catnip_change(self._catnip_redeemed)

    # ------------------------------
    # Animation engine
    # ------------------------------
    def _run_anim(self, frames, msg, speed_ms, on_done):
        self._anim_frames = list(frames or [])
        self._anim_idx = 0

        def _step():
            if self._anim_idx < len(self._anim_frames):
                self._render_frame(self._anim_frames[self._anim_idx], msg)
                self._anim_idx += 1

                # Tiny timing variance (feels less robotic)
                jitter = int(speed_ms * 0.08)
                delay = max(40, speed_ms + random.randint(-jitter, jitter))
                self._after_id = self.parent.after(delay, _step)
            else:
                on_done()

        _step()

    def _return_idle(self):
        self._state = "idle"
        self._idle_idx = 0
        self._render_frame(self.IDLE_FRAMES[0], "...")
        self._idle_loop()

    def _maybe_play_legendary(self) -> bool:
        if random.random() < self._legendary_chance:
            self._cancel()
            self._state = "legendary"
            self._run_anim(
                self.LEGENDARY_FRAMES,
                f"LEGENDARY {self._name.upper()}!! \u2605\u2605\u2605",
                int(300 * self._speed_scale),
                lambda: self._return_idle(),
            )
            return True
        return False

    # ------------------------------
    # Idle loop (expanded)
    # ------------------------------
    def _idle_loop(self):
        self._cancel()
        self._after_id = self.parent.after(random.randint(650, 1500), self._idle_tick)

    def _idle_tick(self):
        if self._state != "idle":
            return

        # Rare legendary pop
        if self._maybe_play_legendary():
            return

        # Seasonal cameo (occasional)
        if self._seasonal_idle_frames and random.random() < 0.12:
            msg = f"spooky {self._name} \U0001f383" if datetime.now().month == 10 else self.MESSAGES["winter"]
            self._cancel()
            self._state = "idle"
            self._run_anim(self._seasonal_idle_frames, msg, int(540 * self._speed_scale),
                           lambda: self._return_idle())
            return

        r = random.random()
        if r < 0.06:
            self._cancel(); self._state = "blink"
            self._run_anim(self.BLINK_FRAMES, self.MESSAGES["blink"], int(300 * self._speed_scale),
                           lambda: self._return_idle())
            return
        if r < 0.10:
            self._cancel(); self._state = "wag"
            self._run_anim(self.WAG_FRAMES, self.MESSAGES["wag"], int(230 * self._speed_scale),
                           lambda: self._return_idle())
            return
        if r < 0.12:
            self._cancel(); self._state = "sleep"
            self._run_anim(self.SLEEP_FRAMES, self.MESSAGES["sleep"], int(640 * self._speed_scale),
                           lambda: self._return_idle())
            return
        if r < 0.14:
            self._cancel(); self._state = "zoomies"
            self._run_anim(self.ZOOMIES_FRAMES, self.MESSAGES["zoomies"], int(200 * self._speed_scale),
                           lambda: self._return_idle())
            return
        if r < 0.24:
            self._cancel(); self._state = "stretch"
            self._run_anim(self.STRETCH_FRAMES, self.MESSAGES["stretch"], int(680 * self._speed_scale),
                           lambda: self._return_idle())
            return

        # Original idle cycle — occasionally show a random thought
        self._idle_idx = (self._idle_idx + 1) % len(self.IDLE_FRAMES)
        thought = random.choice(self.THOUGHTS) if random.random() < 0.15 else "..."
        self._render_frame(self.IDLE_FRAMES[self._idle_idx], thought)
        self._idle_loop()

    # ------------------------------
    # User interactions (kept)
    # ------------------------------
    def _on_frame_click(self, event):
        if self._state != "idle":
            return

        try:
            frame_w = max(self.frame.winfo_width(), 1)
            rel = max(0.0, min(1.0, (event.x_root - self.frame.winfo_rootx()) / frame_w))
            go_frames = self.TREAT_SHORT if rel < 0.30 else (self.TREAT_MEDIUM if rel < 0.65 else self.TREAT_FAR)
            self._cancel()
            self._state = "treat"
            self._total_treats += 1
            self._update_stats()

            if self._maybe_play_legendary():
                return
            if self._maybe_milestone(self._total_treats):
                return

            self._run_anim(go_frames, self.MESSAGES["treat"], int(160 * self._speed_scale),
                           lambda: self._run_anim(self.RUN_BACK, self.MESSAGES["running"], int(160 * self._speed_scale),
                                                  lambda: self._return_idle()))
        except Exception as e:
            print(f"[moveup] Bisa click error: {e}")

    def receive_pet(self):
        if self._state != "idle":
            return

        self._cancel()
        self._state = "pet"
        self._total_pets += 1
        self._update_stats()

        if self._maybe_play_legendary():
            return
        if self._maybe_milestone(self._total_pets):
            return

        self._run_anim(self.PET_FRAMES, self.MESSAGES["pet"], int(260 * self._speed_scale),
                       lambda: self._run_anim(self.HAPPY_FRAMES[:3], self.MESSAGES["pet"], int(240 * self._speed_scale),
                                              lambda: self._return_idle()))

    def throw_treat_at_window_x(self, window_x: int, window_width: int):
        if self._state != "idle":
            return

        try:
            rel = max(0.0, min(1.0, (window_x - (self.frame.winfo_rootx() - self.frame.winfo_toplevel().winfo_rootx())) /
                                    max(self.frame.winfo_width(), 1)))
        except Exception:
            rel = 0.5

        go_frames = self.TREAT_SHORT if rel < 0.30 else (self.TREAT_MEDIUM if rel < 0.65 else self.TREAT_FAR)
        self._cancel()
        self._state = "treat"
        self._total_treats += 1
        self._update_stats()

        if self._maybe_play_legendary():
            return
        if self._maybe_milestone(self._total_treats):
            return

        self._run_anim(go_frames, self.MESSAGES["treat"], int(270 * self._speed_scale),
                       lambda: self._run_anim(self.RUN_BACK, self.MESSAGES["running"], int(270 * self._speed_scale),
                                              lambda: self._return_idle()))

    # ------------------------------
    # New interactions
    # ------------------------------
    def _belly_rub(self):
        if self._state not in ("idle", "pet"):
            return
        self._cancel()
        self._state = "pet"
        self._total_pets += 1
        self._update_stats()
        if self._maybe_play_legendary():
            return
        if self._maybe_milestone(self._total_pets):
            return
        self._run_anim(self.BELLY_FRAMES, self.MESSAGES["belly"],
                       int(240 * self._speed_scale),
                       lambda: self._return_idle())

    def _sneeze(self):
        """Double-click boop -> Bisa sneezes (interrupts pet, doesn't double-count)."""
        if self._state not in ("idle", "pet"):
            return
        was_pet = (self._state == "pet")  # already counted by receive_pet
        self._cancel()
        self._state = "sneeze"
        if not was_pet:
            self._total_pets += 1
            self._update_stats()
        self._run_anim(self.SNEEZE_FRAMES, self.MESSAGES["sneeze"],
                       int(240 * self._speed_scale),
                       lambda: self._return_idle())

    # --- Secret tricks (type while Bisa panel is focused) ---
    TRICKS = {
        "sit":       ("sit",       "SIT_FRAMES"),
        "shake":     ("shake",     "SHAKE_FRAMES"),
        "spin":      ("spin",      "SPIN_FRAMES"),
        "roll":      ("spin",      "SPIN_FRAMES"),      # alias
        "play dead": ("play_dead", "PLAY_DEAD_FRAMES"),
        "dead":      ("play_dead", "PLAY_DEAD_FRAMES"),  # alias
        "daisy":     ("daisy",     "DAISY_FRAMES"),
        "flower":    ("daisy",     "DAISY_FRAMES"),      # alias
        "zoomies":   ("zoomies",   "ZOOMIES_FRAMES"),
        "zoom":      ("zoomies",   "ZOOMIES_FRAMES"),    # alias
        "laser":     ("laser",     "LASER_FRAMES"),
        "box":       ("box",       "BOX_FRAMES"),
        "ryan":      ("bowling",   "BOWLING_FRAMES"),
        "bowling":   ("bowling",   "BOWLING_FRAMES"),   # alias
    }

    def _on_key(self, event):
        """Buffer keypresses on Bisa's frame. If the buffer ends with a trick name, she performs it."""
        if not event.char or not event.char.isprintable():
            return
        self._trick_buffer = (self._trick_buffer + event.char.lower())[-12:]  # keep last 12 chars

        # Special: "help" opens the command popup (not listed as a trick — keeps secrets safe)
        if self._trick_buffer.endswith("help"):
            self._trick_buffer = ""
            self.msg_var.set("need help? \U0001f4cb")
            self._show_help_popup()
            return

        # Special: "debug" plays every animation once as a showcase reel
        if self._trick_buffer.endswith("debug"):
            self._trick_buffer = ""
            self._play_debug_showcase()
            return

        for trigger, (msg_key, frames_attr) in self.TRICKS.items():
            if self._trick_buffer.endswith(trigger):
                self._trick_buffer = ""
                self._do_trick(msg_key, getattr(self, frames_attr))
                return

    def _show_help_popup(self):
        """Show a small themed popup listing Bisa's (non-secret) commands."""
        # Don't stack multiple popups
        if hasattr(self, "_help_popup") and self._help_popup and self._help_popup.winfo_exists():
            self._help_popup.lift()
            return

        popup = tk.Toplevel(self.frame)
        popup.title("Bisa's Commands")
        popup.configure(bg=self._theme_bg)
        popup.resizable(False, False)
        popup.attributes("-topmost", True)
        self._help_popup = popup

        tk.Label(
            popup, text="\u2726 Bisa's Commands \u2726",
            font=("Segoe UI", 11, "bold"),
            bg=self._theme_bg, fg=self._theme_accent,
        ).pack(pady=(10, 4), padx=16)

        tk.Frame(popup, bg=self._theme_border, height=1).pack(fill="x", padx=10)

        commands = [
            ("sit",          "good sit! \U0001f43e"),
            ("shake",        "nice to meet u! \U0001f91d"),
            ("spin / roll",  "wheee~! \U0001f300"),
            ("play dead",    "... \U0001f480 (jk!!)"),
            ("zoomies",      "ZOOMIES!! \u26a1"),
            ("laser",        "RED DOT!! \U0001f534"),
            ("box",          "if I fits, I sits \U0001f4e6"),
        ]

        for cmd, desc in commands:
            row = tk.Frame(popup, bg=self._theme_bg)
            row.pack(fill="x", padx=16, pady=2)
            tk.Label(
                row, text=cmd,
                font=("Courier", 9, "bold"),
                bg=self._theme_bg, fg=self._theme_accent,
                width=13, anchor="w",
            ).pack(side="left")
            tk.Label(
                row, text=desc,
                font=("Segoe UI", 9),
                bg=self._theme_bg, fg=self._theme_msg,
                anchor="w",
            ).pack(side="left")

        tk.Frame(popup, bg=self._theme_border, height=1).pack(fill="x", padx=10, pady=(6, 0))
        tk.Label(
            popup,
            text="type while Bisa's panel is focused  \u2022  click to close",
            font=("Segoe UI", 8),
            bg=self._theme_bg, fg=self._theme_hint,
        ).pack(pady=(3, 10), padx=16)

        popup.bind("<Button-1>", lambda _e: popup.destroy())
        popup.bind("<Escape>", lambda _e: popup.destroy())

        # Position popup just to the right of Bisa's frame
        popup.update_idletasks()
        x = self.frame.winfo_rootx() + self.frame.winfo_width() + 6
        y = self.frame.winfo_rooty()
        popup.geometry(f"+{x}+{y}")

    def _play_debug_showcase(self):
        """Loop through ALL animation sets once — a debug showcase reel. 🎬
        Press Escape to cancel early."""
        if self._state not in ("idle", "pet", "happy"):
            return
        self._cancel()
        self._state = "trick"
        self._debug_cancelled = False

        def _on_escape(_event):
            self._debug_cancelled = True
            self.frame.unbind("<Escape>")
            self._cancel()
            self.msg_var.set("showcase cancelled \U0001f3ac")
            self.parent.after(800, lambda: self._return_idle())

        self.frame.bind("<Escape>", _on_escape)

        # (frames_attr, display_label, speed_ms_base)
        showcase = [
            ("IDLE_FRAMES",       "idle",              300),
            ("PET_FRAMES",        "pet",               260),
            ("BELLY_FRAMES",      "belly rubs",        240),
            ("HAPPY_FRAMES",      "happy",             200),
            ("TREAT_MEDIUM",      "treat chase",       200),
            ("RUN_BACK",          "running back",      200),
            ("SNIFF_FRAMES",      "sniff sniff",       260),
            ("ALERT_FRAMES",      "alert!",            280),
            ("LOAD_FRAMES",       "data loaded",       280),
            ("KUNTAL_FRAMES",     "priority item",     300),
            ("SUCCESS_FRAMES",    "success",           250),
            ("WARNING_FRAMES",    "warning",           320),
            ("CONFUSED_FRAMES",   "confused",          280),
            ("EXCLUDED_FRAMES",   "excluded",          320),
            ("CLEARED_FRAMES",    "cleared",           320),
            ("WAG_FRAMES",        "tail wag",          230),
            ("BLINK_FRAMES",      "blink",             300),
            ("SLEEP_FRAMES",      "sleep",             400),
            ("STRETCH_FRAMES",    "stretch",           400),
            ("ZOOMIES_FRAMES",    "zoomies",           200),
            ("SIT_FRAMES",        "sit",               240),
            ("SHAKE_FRAMES",      "shake",             240),
            ("SPIN_FRAMES",       "spin",              240),
            ("PLAY_DEAD_FRAMES",  "play dead",         240),
            ("SNEEZE_FRAMES",     "sneeze",            240),
            ("BOX_FRAMES",        "box",               240),
            ("BOWLING_FRAMES",    "bowling",           240),
            ("LASER_FRAMES",      "laser",             240),
            ("DAISY_FRAMES",      "daisy",             240),
            ("LEGENDARY_FRAMES",  "legendary",         300),
            ("CATNIP_DISCOVER",   "catnip: discover",  150),
            ("CATNIP_HIGH",       "catnip: high",      220),
            ("CATNIP_CHILL",      "catnip: chill",     420),
            ("CATNIP_RECOVER",    "catnip: recover",   320),
            ("HALLOWEEN_FRAMES",  "halloween",         350),
            ("WINTER_FRAMES",     "winter",            350),
        ]

        total = len(showcase)

        def _play_next(idx):
            if self._debug_cancelled:
                return
            if idx >= total:
                self.frame.unbind("<Escape>")
                self.msg_var.set("showcase done!! \U0001f3ac\u2728")
                self.parent.after(1500, lambda: self._return_idle())
                return
            frames_attr, label, speed = showcase[idx]
            frames = getattr(self, frames_attr, None)
            if not frames:
                _play_next(idx + 1)
                return
            self._run_anim(
                frames,
                f"[\U0001f3ac {idx + 1}/{total}] {label}  (Esc=stop)",
                int(speed * self._speed_scale),
                lambda i=idx: _play_next(i + 1),
            )

        self.msg_var.set("\U0001f3ac debug showcase starting...  (Esc=stop)")
        self.parent.after(800, lambda: _play_next(0))

    # Primary tricks only (no aliases) — used by the random trick button
    _RANDOM_TRICKS = [
        ("sit",       "SIT_FRAMES"),
        ("shake",     "SHAKE_FRAMES"),
        ("spin",      "SPIN_FRAMES"),
        ("play_dead", "PLAY_DEAD_FRAMES"),
        ("zoomies",   "ZOOMIES_FRAMES"),
        ("laser",     "LASER_FRAMES"),
        ("box",       "BOX_FRAMES"),
        ("daisy",     "DAISY_FRAMES"),
        ("bowling",   "BOWLING_FRAMES"),
    ]

    def _play_random_trick(self):
        """Pick a random trick and play it — triggered by the 'click for a trick' label."""
        msg_key, frames_attr = random.choice(self._RANDOM_TRICKS)
        self._do_trick(msg_key, getattr(self, frames_attr))

    def _do_trick(self, msg_key: str, frames: list):
        if self._state not in ("idle", "pet", "happy"):
            return
        self._cancel()
        self._state = "trick"
        self._total_pets += 1
        self._update_stats()
        self._run_anim(frames, self.MESSAGES[msg_key],
                       int(240 * self._speed_scale),
                       lambda: self._run_anim(self.HAPPY_FRAMES[:4],
                                              "good kitty!! \u2728",
                                              int(200 * self._speed_scale),
                                              lambda: self._return_idle()))

    def _on_hover(self, event=None):
        if self._state != "idle":
            return
        if random.random() < 0.28:
            self._cancel()
            self._state = "sniff"
            self._run_anim(self.SNIFF_FRAMES, self.MESSAGES["sniff"],
                           int(260 * self._speed_scale),
                           lambda: self._return_idle())

    def _maybe_milestone(self, count: int) -> bool:
        """Bisa celebrates at fixed milestones OR every ~80 interactions (+/-20)."""
        self._interactions_since_milestone += 1
        fixed_milestones = {10, 25, 50, 100, 200, 500, 1000}
        interval_hit = self._interactions_since_milestone >= self._next_milestone_interval
        if count not in fixed_milestones and not interval_hit:
            return False
        # Reset interval counter and pick next random threshold
        self._interactions_since_milestone = 0
        self._next_milestone_interval = random.randint(60, 100)
        stars = "\u2b50" * min(5, (count // 100) + 1)
        msg = f"{stars} {count} total!!"
        self._run_anim(
            self.KUNTAL_FRAMES,
            msg,
            int(220 * self._speed_scale),
            lambda: self._run_anim(self.HAPPY_FRAMES[:3], msg,
                                   int(200 * self._speed_scale),
                                   lambda: self._return_idle()),
        )
        return True

    def greet_startup(self):
        """Bisa greets the user based on time of day when the app launches."""
        hour = datetime.now().hour
        if hour < 6:
            msg, frames = "up late?? \U0001f319", self.SLEEP_FRAMES
        elif hour < 12:
            msg, frames = "good morning!! \u2600\ufe0f", self.LOAD_FRAMES[:3]
        elif hour < 17:
            msg, frames = "good afternoon~ \U0001f324\ufe0f", self.HAPPY_FRAMES[:3]
        elif hour < 21:
            msg, frames = "good evening! \U0001f306", self.WAG_FRAMES
        else:
            msg, frames = "working late? \U0001f319", self.BLINK_FRAMES
        self._cancel()
        self._state = "idle"
        self._run_anim(frames, msg, int(300 * self._speed_scale),
                       lambda: self._return_idle())

    # ------------------------------
    # Reactions (kept)
    # ------------------------------
    def celebrate(self):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "happy"

        if self._maybe_play_legendary():
            return

        self._run_anim(self.HAPPY_FRAMES, self.MESSAGES["happy"], int(480 * self._speed_scale),
                       lambda: self._return_idle())

    def react_data_loaded(self, row_count: int = 0):
        if self._state not in ("idle", "happy"):
            return
        self._cancel()
        self._state = "loaded"
        self._run_anim(self.LOAD_FRAMES, self.MESSAGES["loaded"], int(420 * self._speed_scale),
                       lambda: self._return_idle())

    def react_moveups(self, count: int):
        """Bisa celebrates when SKUs are detected as moved to Sales Floor since last load."""
        if self._state not in ("idle", "happy", "pet"):
            return
        self._cancel()
        self._state = "moveup"
        msg = f"{count} SKU{'s' if count != 1 else ''} moved!! \U0001f4e6"
        self._run_anim(
            self.ZOOMIES_FRAMES, msg, int(200 * self._speed_scale),
            lambda: self._run_anim(self.HAPPY_FRAMES[:3], msg, int(220 * self._speed_scale),
                                   lambda: self._return_idle()),
        )

    def react_excluded(self, count: int = 1):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "excluded"
        self._run_anim(self.EXCLUDED_FRAMES, self.MESSAGES["excluded"], int(520 * self._speed_scale),
                       lambda: self._return_idle())

    def react_restored(self, count: int = 1):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "happy"
        self._run_anim(self.HAPPY_FRAMES[:4], self.MESSAGES["restored"], int(500 * self._speed_scale),
                       lambda: self._return_idle())

    def react_row_selected(self):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "alert"
        self._run_anim(self.ALERT_FRAMES, self.MESSAGES["alert"], int(380 * self._speed_scale),
                       lambda: self._run_anim(self.SNIFF_FRAMES, self.MESSAGES["sniff"], int(460 * self._speed_scale),
                                              lambda: self._return_idle()))

    def react_kuntal(self, count: int = 1):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "kuntal"
        self._run_anim(self.KUNTAL_FRAMES, self.MESSAGES["kuntal"], int(440 * self._speed_scale),
                       lambda: self._return_idle())

    def react_cleared(self):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "cleared"
        self._run_anim(self.CLEARED_FRAMES, self.MESSAGES["cleared"], int(500 * self._speed_scale),
                       lambda: self._return_idle())

    # ------------------------------
    # New contextual reactions
    # ------------------------------
    def react_success(self, msg: str = "nice!! \u2705"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "success"
        frames = self.SUCCESS_FRAMES + self.WAG_FRAMES
        self._run_anim(frames, msg, int(250 * self._speed_scale), lambda: self._return_idle())

    def react_warning(self, msg: str = "uh oh\u2026 \u26a0\ufe0f"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "warning"
        self._run_anim(self.WARNING_FRAMES, msg, int(320 * self._speed_scale), lambda: self._return_idle())

    def react_error(self, msg: str = "nope\u2026 \U0001f4a5"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "error"
        self._run_anim(self.CONFUSED_FRAMES, msg, int(300 * self._speed_scale), lambda: self._return_idle())
