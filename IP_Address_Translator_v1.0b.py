# Author: Abdallah Hamdy
# Purpose: Converting subnets and IP addresses into address objects.

# This version is a draft version, and it is not the final version. so do not judge the code quality.


import os
import re
import sys
import copy
import time
import socket
import datetime
import ipaddress
import threading
import webbrowser
from base64 import b64decode

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.constants import *

import paramiko
import xlsxwriter
import dns.resolver
import pandas as pd
from cryptography.fernet import Fernet
from fortigate_api import FortiGateAPI

# Show the splash screen
try:
    if getattr(sys, 'frozen', False):
        import pyi_splash  # type: ignore

        pyi_splash.update_text('UI Loaded ...')
except:
    pass

# Constants
NAME = "IP Address Translator"
SHORT_NAME = "IAT"
VERSION = "1.0b"
AUTHOR = "xAbdalla"
GITHUB = "https://github.com/xAbdalla"

FULL_NAME = f"{NAME} ({SHORT_NAME})"
TITLE = f"{FULL_NAME} v{VERSION} by {AUTHOR}"
TITLE_LABEL = f"{NAME} v{VERSION}"

MASTER_KEY = b"TLAQL1vKnWRoleHlIpvrUnpkATe1GBlFChUq78sqB6I="  # Base Key to encrypt/decrypt the KEY used in the app.
cipher_suite = Fernet(MASTER_KEY)
while True:
    if os.path.exists("settings.cfg"):
        try:
            with open("settings.cfg", "r") as f:
                encrypted_key = f.read()
            KEY = cipher_suite.decrypt(encrypted_key.encode('utf-8'))  # Import saved KEY from the previous run
            break
        except:
            # If you can not decrypt the KEY, rename the settings file to invalid
            os.rename('settings.cfg', f'{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}_settings.cfg.bkp')
    else:
        #  If the settings file does not exist, generate a new KEY and save it
        KEY = Fernet.generate_key()
        cipher_text = cipher_suite.encrypt(KEY)
        with open('settings.cfg', 'w') as f:
            f.write(cipher_text.decode('utf-8'))
        break

# Variables Initialization
input_file = ""
output_file = ""
ref_file = ""

inputs = {}  # {sheet1: [[subnet1, type1], [subnet2, type2], ... etc], sheet2: [[subnet1, type1], ...etc]}
inputs_bkp = {}
input_invalids = []  # [[sheet1, subnet1, row1], [sheet2, subnet2, row2], ... etc]
inputs_no = 0  # Number of valid subnets in the input file

outputs = {}

refs = {}  # {sheet1: [[tenant1, name1, subnet1, type1], [tenant2, name2, subnet2, type2],... etc], ...etc}
refs_bkp = {}
ref_invalids = []  # [[sheet1, subnet1, row1], [sheet2, subnet2, row2], ... etc]
ref_no = 0  # Number of valid subnets in the reference file

ssh_panorama = None
panorama_ip = ""
panorama_username = ""
panorama_password = ""
panorama_vsys = ""
panorama_addresses = []  # [[vsys1, name1, subnet1, type1], [vsys2, name2, subnet2, type2], ... etc]
panorama_addresses_bkp = []  # This is workaround, because I did not want to change the code too much (will be fixed in the next version)

forti_api = None
forti_ip = ""
forti_port = "443"
forti_username = ""
forti_password = ""
forti_vdom = ""
forti_addresses = []  # [[vdom1, name1, subnet1, type1], [vdom2, name2, subnet2, type2], ... etc]
forti_addresses_bkp = []  # This is workaround, because I did not want to change the code too much (will be fixed in the next version)

ssh_aci = None
aci_ip = ""
aci_username = ""
aci_password = ""
aci_class = "fvSubnet"
aci_addresses = []  # [[tenant1, name1, subnet1, type1], [tenant2, name2, subnet2, type2], ... etc]
aci_addresses_bkp = []  # This is workaround, because I did not want to change the code too much (will be fixed in the next version)

dns_resolver = None
dns_servers = dns.resolver.Resolver().nameservers
resolvers = []

start_thread = None
start_flag = False
stop_flag = False
flags = [False for _ in range(5)]  # [File, Panorama, ACI]

separators = [",", ";", "\n", "\r", "\r\n", "\n\r"]

# Create the main window
gui = tk.Tk()
gui.title(TITLE)
gui.focus_force()
gui.grab_set()

# Set the window size
WIDTH = 550
HEIGHT = 470

# Get the screen size
SCREEN_WIDTH = gui.winfo_screenwidth()
SCREEN_HEIGHT = gui.winfo_screenheight()

# Calculate the x and y coordinates
X = int((SCREEN_WIDTH / 2) - (WIDTH / 2))
Y = int((SCREEN_HEIGHT / 2) - (HEIGHT / 2) - 100)

# Set the window geometry
gui.geometry(f"{WIDTH}x{HEIGHT}+{X}+{Y}")
gui.resizable(False, False)


def close_gui_window():
    global gui

    start_end()
    disconnect_from_panorama(False)
    disconnect_from_forti(False)
    disconnect_from_aci(False)
    gui.destroy()


def write_logs(*arg):
    global logs
    log = logs.get()
    with open("logs.txt", "a") as f:
        if log:
            f.write(f"{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - {log}\n")
        else:
            # If the log is empty, write a line separator
            f.write(f"{"=" * 50}\n")


# Check if the log file exists and big delete it
# if os.path.exists("logs.txt"):
#     if os.path.getsize("logs.txt") > int(10 * 1024 * 1024):
#         os.remove("logs.txt")

# Logs Variable
logs = tk.StringVar()
logs.trace_add("write", write_logs)
logs.set("INFO - Application started.")

# Icons Data
icon_data = (
    "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")
github_icon_data = (
    "iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAQAAABecRxxAAAzUUlEQVR42u2dd5xV1bm/n5mh16F3EAQpdhBULKBiFzvWiCmK"
    "xiSaYqIm915Ju5eY3F+CSUjwpqmxISYm2HtHBETAglKUIr0MXRhm5vcHZdqZM+fs9t1r7/d5Px/lzNnn7O9611rvWWvtVcAw"
    "jNRSoBZgiGhC02qvt7NbLcmIHgsAyaMJHehKR9rShmKKaUMxLWlFE5rSgoa0yfLZcjazm+1sZzclbKWEEjZRQgkbWMNq1lqY"
    "SBYWAFynLT3oSU+604OedKALrUO930bWsIblLGMFy1nKcjarXWB4xwKAe3TgEPrRl770oy+t1HLYxEIWsYhPWMRCNqrlGPlg"
    "AcANijmYQxnEoQyhi1pMVjbxIR/wIR8wj7VqMUZ9WACIM133VfkhDHQyp1Yxm9nMZiar1VKMzLhYrJJOS4ZyPMcxlE5qKYHx"
    "OTOZznRms0MtxaiKBYD40IsRnMBxHEqRWkpo7GEu03mLV1mplmKABYA40JUTGMWJDFILiZQlvMkbPMdnaiHpxgKAjjacxhmc"
    "QS+1ECkLeY7neJmtaiHpxAJA9BRxLGdxOkMT3NTPl1Km8xzPMpsKtZR0YQEgSppxGucxOuYP8pSs4xmm8TTb1ELSggWAaOjM"
    "hVzAKTRWC3GCHTzPv/kXG9RCko8FgLBpzzmM4SwaqIU4Rxlv8yiP2ByCMLEAEB6duJwxDKdQLcRpyniZR5lqU4zDwQJAGDTl"
    "PMZyJg3VQhLDbp7jUR5ju1pI0rAAECyFnM41XEhztZBEUsJj3Mfr9qQgOCwABEd3ruYGeqtlJJ7lPMgfWKqWkQwsAARBEy7i"
    "a5xivf3IKONp/syTlKqFuI4FAL9043puooNaRipZzb1MYplahpFOChnFFEqpMBNaGdMYZT9kXjHHeaMVX+WbHKyWYezjQ+7m"
    "fltqnD8WAPKnK+O4OevWmoaCzdzLL1mhlmEkmeOZwh55s9esLtvF3zhSXUiMZHIi0+QF3CwXe4PR6sJiJIlCRjNDXqzN8g0C"
    "1sE1fFPEWBbIi7OZF5vDJRYEDO8UMJr35MXYzI/NY4y6GBluMoqZ8uJrFoS9xSnqwmS4xXBekhdbsyDteY5RFyrDDQ5niry4"
    "moUTBOwRoZGVgUyhXF5QzcKyMqbQV13I4oSNkVbSjZ9zja3oSzyl3MOdtt/gXmxj6r005Xs8wjALiCmgiGGMA2ZRppaixwo8"
    "FHApd3GQWoYRMQv5EY+qRaixAHAMv+NYtQhDxIt8i4/UIpSku8fbhom8bdU/xZzGXCbSSi1DR3rHAAoYy785NeUh0CjiWK5l"
    "PfPUQjSktQtwNH+wX36jCi9yE5+oRURPGlsATfkP7k35mbxGTfpwI815PW1PBtLXAhjBZPqrRRgx5X3GMV0tIkrS1QNuz328"
    "YtXfqJPDeIPf0lItIzrS1AI4l3voqhZhOMBSruMFtYhoSEsLoJjJPGHV38iJXjzH5HS0A9LRAjiHe+imFmE4xlK+xotqEWGT"
    "/BZAS/7Gk1b9jbzpxfNMpIlaRrgkvQVwDA/STy3CcJgPuTLJk4SSPA+ggFt4mI5qGYbTdOArbGOGWkZYJLcF0Im/cZZahJEQ"
    "Hue6ZO4gkNQAcAb30lktwkgQK/gSr6pFBE8SBwGbMJFnrPobgdKdl5lIQ7WMoEleC2AgD3KUWoSRUN7hKharRQRJ0loA1zPb"
    "qr8RGsOYzaVqEUGSpKcAjfkD45PXSDNiRRPG0JSXqVALCYbkdAG6MpXj1SKMlPAsV7FRLSIIkhIAhjOVLmoRRopYzEXMV4vw"
    "TzLGAMbxslV/I1IOZnoSjh11fwygMX/gzgSkw3CNRlxKW16gXC3ED653Aaznb2h5hctYpxbhHbcDgPX8DT3LuJjZahFecXkM"
    "4Frr+RsxoCevcr5ahFfc7Tvfwh9ooBZhGEAjLmMDM9UyvOBmAChiEv/hePfFSBKFnENbnnNvepCLlag5j3CuWoRh1GIq1/CF"
    "WkR+uBcAOvMEQ9QiDCMjL3MxJWoR+eBaAOjL0/RVizCMOvmAc1imFpE7bj0FOI7pVv2NWHMob3O0WkTuuBQALuIl2qtFGEY9"
    "dOE1zlaLyBV3AsC3mEpTtQjDyIEWPM6VahG54cpjwFv5tXPjFUZ6KeIiVjBHLSMXoS5wG3epJRhGXhRwPht5Ry2jPlwIAOP5"
    "mVqCYeRNAWexhbfVMrIT/wDwM/5TLcEwPFHAmZTzmlpGNuIdAAr4Dd9XizAMH5xC0zgfNR7nAFDAb/mWWoRh+OREmsU3BMQ3"
    "ABTxJ8apRRhGAJxAR55Wi8hMXANAEX/lWrUIwwiIoXThqTiuFYxnACjkfr6kFmEYATKEdnFsBcQxABQwia+oRRhGwAyjiJfV"
    "ImoSxwAwgW+rJRhGCIxgB2+pRVQnfgHgDu5USzCMkBjFqnhtIBq3APB1/p9agmGERgHn8gnvq2VUFRQnruY+h9YnGoYXSrmI"
    "J9Ui9hOnADCax+xsXyMF7OSsuEwQjk8AOIWnaKIWYRiRsIVT4zEWEJcAMIwXaKkWYRiRsZ4RfKgWEZcA0JfpttmXkTKWchyr"
    "1SLiMOTWmset+hupoxdP0FwtQh8AGvIYh6pFGIaAIdyrroH6eQCTGKOWYBgiBtGIF5UC1AHgDm4TKzAMJSexWvk8QDsIeAlT"
    "1E0gwxBTytm6VoAyAAzhVf0giGHI2chwPtbcWhcAujGDbrK7G0ac+JRjWae4saoB3pKnrPobxj568w8aK26sGQQs5B+cKLmz"
    "YcSTnvTg8ehvqwkAd3K95L6GEV+OYi2zor6pYgxgFM/IHz8aRvwoZWTUOwZFHwB6Mtsm/hpGRpYzmPVR3jDqQcDGTLXqbxh1"
    "0IOHo20dR90U/z3nR3xHw3CJPlTwSnS3i7YLcBUPRHo/w3CPcs6L7gSBKAPA4bxNswjvZxhusokhfBrNraIbA2jJFKv+hpED"
    "bXgkqmlBUY0BFPAAJ0d0L8NwnW60i2bn4KgCwHf4TkR3MowkMJSPozg/IJoxgEHMomkkdzKMpLCZI1gW9k2iGANozINW/Q0j"
    "T1rz9/DrZ4MIEvJTjozgLvkwlddZzkrWAVtoRRta0Z1e9KQfR1GslmeEyDrmspBlLONzNrOFzbSmgE50owdncoZaXjVO4tth"
    "H5UXfhdgBC/FbNeftXRjT5b3e3EkwxjBUM0CTSMEdvA2rzKTuazMclUbVsUsz3cxlPlh3iDsANCaufQK+R758ltuzum6phzH"
    "SM5lcExOTzDyp5zpPMWrzGR3Ttf/kwvVkmswj2HsUovwzt+piJ0NyzMNHRnLNL6Q6zbLx3byPLfQNc+8vlSuu7b9Ul2JvXOZ"
    "3Hm1baHH3/PWfJVXKJfrN6vP9vA0V3mcdNaEErn+mlbGSHVF9kZXNsidV9vG+0pTd27jY3kazOqyDxjvs8v5F3kaatty2qgr"
    "c/4U8JzccZnskADSdiIPsVueErOqtpM/MySAvD1NnpJM9nd1dc6fb8qdlsneDSx9XfkZa+TpMaugghX8kA4B5Wshq+XpyWSX"
    "aaqxV3qwRe6yTPaTQFPZiDFMl6cp3TaLsTQMNFf/Jk9TJlsXWIiLhMflDstsx4WQ1hN5QZ6udNobnBZCfsZx6LqCCu4Lu9IG"
    "RxwfplRQwbrQFj+dyDR56tJlz4cSzAFax3Z8Z3RYFTZYWrFC7qrMdn+o6T6eafaYMAIrZxpDQ83JV+RpzGxLaRl0UsOYpDsh"
    "tmf+PBXqt09nNEN5Tp3IhPM4RzCamaHeI5KV+B7oyX+rJdTPcMrkkTKzldE2Eg+czhx5WpNp0yM6T+oweUrrLsPhtn1804D3"
    "5E6qy+ZE5oUCxrBYnt5k2WeMjWxNRgFr5emty2YFO44VdBfg9tgt/a3kzcjuVMGjDOQGzXmvCWQDt9Of+6iI6H4VvK1Ocp0M"
    "4etqCXXTj53yCFm3XRW5P9ryG0rl6XbbdjGBVpHn3G3ydNdtJXSJ3B85UcDzcudks54Sr/SPuVfibS8zSJJrJ8hTns0eCi6h"
    "Qfaq4n3sxwp6iO5cwFjuoqM4/dvZyjY2s5lyYAtlwE6+AKAZjYGGtAAa0JJiWtJSvo3b53yPR0T3bkJJzLYGqc6pvBzMFwUX"
    "AJqzgO4id+TCFC4X3r2Yn3FjBHswb+Fz1rGWNaxnHWtYy0a2smVftc+PIlpRTAva0omOdKADnelAe3rQPPR0lPJbxrM19PvU"
    "zZsMF969PuYzOOuuVjkT3J6Ad8S6+hPhM4BMlPBN/sIkjg34e9eymGUsYxlLWcoySgL87jI2sSnjO+3oSU960ZOe9KQv7QJO"
    "1Wt8I4otsbPyXqwDwOGMY1IQXxRUC6A3H9JE6JD6OZ9pagkUch130drXd+ziEz7mEz7mYz6po4JGTTv605/+HEJ/+tLI13dt"
    "4LvcT1Tj/XXz9WAqWGhs4BA2qkVU8g/5wEh91lvton104wkP6j/nSf6bMfSL2QartSliIFcwgWc9Lat9jE7qBOzjJHmJrc9+"
    "p3ZRJafKnVGfbY3Vxp5j2ZhjtZ/KbZwhH0D0ShfO5g7+yaqcUrtOOkpTkzbyMluflXKY2kl7KWCW3Bn12Qy1k2rQiceyZOwH"
    "TGYsh6pFBkhXxjCRN9hVZ6qn5b2BZ9jEdUlbpcVkzcLVckfUb39VOykDV7GumsZdvMp/cZL88VuYNGck43mjxoLbNVyqFpaB"
    "Z+Wltn47Ve0kaOTErPf/UrspIx2ZQgUVLGYyY3wODrpFM0YxgVmUUcGUmO5180d5qa3f3tOPCd0qd0Iu9hW1m+rkONqrJQjp"
    "GPiD0eD4kbzU5mJf0jqpTY7DWWoLY+soI9lcIy+1udhn/h6/+21A3ObIjuVL1QIM51iuFpATvbjRz8f9PRzrzKIIpoX6p4Jm"
    "+2a9G0au9GGxWkJOrKeP90nT/loAP3Si+sMmq/5G3qxSC8iR9nzL+4f9tAB68kmsV0xVsoSD1RIMB/nCkfJdQh+vk8L9tADu"
    "dMQ9BLpExkgPm9UCcqSY73r9qPcA0Jex6nTnzBa1AMNJStQCcubbXh8mew8APwxwKXHYxGPNnOEaJWoBOdOCb3v7oNcAcJB6"
    "AkJebFMLMJxEuSFJvnzL2wN5rwHg9oAPZAyXUrUAw0kC2XMnIlpxs5ePeQsA3fmyOr15kf92WIbhWrn5DsX5f8hbAPiBM+P/"
    "eylTCzCcxK1y09rLiQFeAkAHrlOnNU/0G0wZLuJWCwC+k/9Sci8B4JvOrVi3AGB4wbVy0yH/rnn+AaAZN6nTmTeuBSwjHjRT"
    "C8ibW/Pdej7/AHCdg+vX3VixYMQN9wJAHy7K7wP5BoAibw8bxLiXkUYccLHc/CC/y/MNAGOcXFZjLQDDCy52HYcyIp/L8w0A"
    "31anzxPFagGGkxSrBXjiO/lcnN9y4OOYrk6dJ1bFbstpI/4UsSuC0xyDp5x+LMn14vxaAD42HpDSQb93quEc7Z2s/lCYz3O6"
    "fFoAXfjM57lvOjqwXi3BcIwjmKuW4JESeuS6AC6fX8ZxzlZ/6KIWYDiHu2WmmGtyvTT3ANDI3+6jYlx8dmFo6aMW4INv5tq2"
    "zz0AXEhndap84HJmGhpc/tEYxEm5XZh7ALhBnSZfuJyZhga3fzRyrK+5BoB+nKJOkS/6qwUYznGIWoAvLsltyn6uAWCczyNE"
    "1ByhFmA4RmPHA0Bjrs3lstyqdSNWxPQE19zpwmq1BMMhBjNbLcEnHzOw/gXNubUALnG++lsbwMgP98tL/1xWBeQWAL6sTksA"
    "DFMLMJwiCeUlh05ALgGgWyIO1z5eLcBwiiSUlzG0rO+SXALAWEfnRFdnuK0HMHKmOYepJQSSiovruyS3AJAEijlSLcFwhpMd"
    "OvcqG/V2AuoPAMcxQJ2KgDhbLcBwhqSUlZH1TWeqPwDkvKwg9iQlU43wSUpZKeDq+i7ITgM+p6M6FQFRRkc2qkUYDtCPT9QS"
    "AmMBA7O9XV8L4PTEVH8oSsTTDCN8kvL7DzAg+9hXfQHgCrX+QKl3TNQwgAvVAgLlymxvZu8CNGE1rdX6A2QnXdisFmHEnG4s"
    "S9Qj42UcVPeU4OwJPTdR1R+aWhvAqJexiar+0JPj6n4ze1IvV2sPnCFqAUbsOVotIHCy1ONsXYAmrK1/KqFTzOHkXDdLNFJL"
    "O6bTTy0iUFbQs65OQLYWwJkJq/4rucCqv1EvGzibdWoRgdKdY+p6K1sAuFCtO1C2cg7L1SIMJ1jMpexSiwiUPI8MBWjAOioS"
    "Y+Vcos4DwynGystskOZhYtNpctFB2k/U5clwjsnyUhukHZo5kXV3AS5Q+z9AnmG8WoLhHDczQy0hQC7M/Oe6nwIsobdac0B8"
    "xhBbA2B4oDvvJmAzvL28nXmLk7paAIMSU/33cLVVf8MTK7iWerfVdIRhmVf11BUAzlXrDYwf85ZaguEsT/N7tYSAKOSMfC5/"
    "RT5oEYy9lojtzAwdjXlPXoqDsQczJS/zGEAxa2mo9nwAbOFwlqlFGI5zOLMcPhe7ko10Yk/NP2buApyeiOoPP7Dqb/hmPj9T"
    "SwiEtpkWBWUOAHn1FmLLK9yjlmAkgv9x/pSgvZxe+0+ZA0ASds7ZwfWJGcE1tOzha5SqRQTAqNp/yhQADk7EI8Afs0gtwUgM"
    "c5molhAAw2rv75EpAJyew1fFnYWJyDAjPvyElWoJvmlQ+7TATAEgCR2AWxK2mstQs5Xb1RICYFT9lxSyXv7E0q89pvazkUAK"
    "eF1esv3ah7UTVZMjeU/taZ+UMpDFahFGAjmGd+o9SSPudGVV1Ze1uwAnqxX6ZrJVfyMUZiWgbXlC9Ze1A8BJaoU+2c7P1RKM"
    "xPKj2nPpHKNG/a4dAE7I8Yviyq9ZrZZgJJZP+Ktagk9qBICaPRrXT0XbSi82qUUYCaYni5yeKF9Gu6qH49RsAbg+AjDJqr8R"
    "Kssyr6pzhiKGV31ZMwAMz+Or4scX/EYtwUg8EyhXS/BFtU5+zQBwrFqdL/5i/X8jdBbwuFqCL6rV8epjAC0pcfhctHIOsQeA"
    "RgQcx3S1BB9spm1lG6Z6dR/qcPWHp636G5HwNrPUEnzQmv6VL6pXeLc7AJPUAozU4HZZq1LPqweAYWplPljCM2oJRmp4mA1q"
    "CT6oMwAck+cXxYl7HB+bNVxiJ/eqJfhgaOU/qw4CdmCtWplnyunFCrUII0W4vGhuFy3373BUtQVwtFqXD16w6m9EylzmqyV4"
    "pjED9/8zKQHgfrUAI3X8XS3AB0ft/0fVAHCEWpVntvFPtQQjdTzo8KjTUfv/UTUADFar8sxTbFdLMFLHCoenAx21/x+VAaAZ"
    "/dSqPPNvtQAjlTyuFuCZI/f/ozIADHT2FL1SnlJLMFLJ42oBnmlL173/qAwAg9SaPPOKLQE2JCyqvcmmM+yr75UBYIBakWee"
    "UAswUss0tQDP7HsQmIQWwAtqAUZqeVEtwDOJCQBr+EgtwUgtbzp7AE2NLkAj+qgVeeRFOwLUkLGDt9USPFIjAPSjgVqRR15S"
    "CzBSjavlrwPtoDIAHKzW45nX1QKMVONu+TsYKgNAX7Uaj2xioVqCkWpmOTshuA+43wKYZSMAhpStLFBL8Ei1FoCrAeAdtQAj"
    "9cxUC/BIIgKAy5szGskgAQGgAb3UajwyTy3ASD2ulsEqAaCro6ed7eQztQQj9bg6BtCFRvsDQHe1Fo8scHYE1kgO61ivluCJ"
    "QrruDwA91Fo8YpOAjTjgahugx/4A0E2txCOuOt5IFq6Wwx6utwCWqAUYBvCpWoBHurs+BrBcLcAwcLcc9qh8CuAmdhaAEQdc"
    "DQDd9geAzmolnqjgc7UEwwCWqQV4pNP+ANBBrcQTa53djMFIFp87uiKl494A0IyWaiWecPckQyNZ7GKLWoIn9gWATmodHtmo"
    "FmAY+3BzX+pWNC0EOqp1eMRNpxtJxNWy2NHlAFCiFmAY+3C1NdqxEPbuDeYgrjrdSB4lagEeaVsIFKtVeMQOBDXigqtlsdjl"
    "AFCqFmAY+9itFuCR4kKgtVqFRywAGHHBAoAAV51uJA9Xf4xaWxfAMPzj6o9RcSHQSq3CI7YbkBEXytQCPNK6EGiuVuERVw8z"
    "M5KHm3tqQrNCoKlahUdcdbqRPFwti033LgZyE1edbiQPV8ui0y0A6wIYccHVAGAtAMMIAFfLorUADCMAXA0ATQuBRmoVHnFV"
    "t5E8XP0xalxIAUVqFR5x1elG8nD1x6hBobPVHxqrBRjGPlwNAEUuBwBX1zAYyaNYLcAjFgAMIwBcLYtOdwFcdbqRPFwti9YC"
    "MIwAcHNjfWhQ6PCaOgsARjwodHZFbVmhsyuZ3R14MZJGKwrUEjxSWujwthotHO6+GEnC3bZoaSFlznYCCpzteRnJwukA4PLW"
    "Wm4eamokjfZqAZ4pLQT2qFV4pqtagGEA3dQCPLPb7RaABQAjDrhbDh3vArjreCNJdFEL8IzjAcBdxxtJwroAIqwFYMQBd8th"
    "aSGwS63CM+463kgS7rZEdxcCW9QqPGMBwNBT4HAA2FwIbFar8Ex3Z6dgGsmhA03UEjyzxe0A0JTuaglG6jlELcAHJYVAiVqF"
    "D1x2vpEM+qkF+KDE7RaABQBDj8sBwPExALedbyQDl3+ELAAYhk9cLoPOBwCXo6+RBAroq5bgg82uDwL2cfZQJiMZdHf2bE1I"
    "wCBgA2sDGFIGqQX4YkshsEGtwheD1QKMVDNELcAXGwqBVWoVvrAAYChxOQDsYW0hsM7h9YAWAAwtLpe/1ZQXAhWsVivxwWAK"
    "1RKM1NKWXmoJPljJvsrjcieghdPPYQ23GeL0crRV+wPASrUSX7jcCDPcxuURgIS0AFzPBMNl3P7xWZ2MAHCyWoCRUgo4SS3B"
    "FyuT0gUoVkswUskAOqsl+CIhXYAihqslGKlkpFqATxIyCOh+RhhuMkItwCcHWgBL1Ep84npGGG7i9gjAdtbuDwBbWKtW44vB"
    "tFJLMFLHQMd3pV5EBQdm0S1Sq/FFA3sSYETOSLUAnyyGygCwWK3GJ+epBRipY7RagE8WQXICwPlOT8k03KMFp6gl+KRaC8Dt"
    "LgB04Ri1BCNVnOnwcSB7qdYCcD0AwPlqAUaqcL+8LQIONJzbsV6txyfzOFItwUgNRaymvVqEL3bRnLLKFsAGNqkV+eQIeqsl"
    "GKnhRMerPyyhDKiymYb7nYBL1AKM1HCxWoBv9g37VwaAD9WKfPNltQAjJTTgcrUE33yw93+VAWC+WpFvDuVotQQjFZxNJ7UE"
    "3+yr75UBYJ5aUQBcoxZgpIIklLO5e/9XOX2mg+PrAQDW0t3pHY4NF2jNKpqqRfhkNy3ZDVVbAOtYo1blm46coZZgJJ4rnK/+"
    "8OHe6k+1LbWT0An4qlqAkXiuVQsIgAMjfkkLABfYbAAjVIZwvFpCAMzd/4+kBYAiblRLMBLNd9QCAuFAXa+6hu4o5qh1BcAm"
    "erBdLcJIKF35lEZqEQHQef+IX9UWwEeJGEFvw5fUEozEclMiqv/aygH/qgFgVwJmAwLcbHsDGKHQhHFqCYHwbuU/qx+sOV2t"
    "LBAGcbZagpFIxtJBLSEQqtTz6gHgbbWygPiJtQGMwGnI7WoJAVFnAEhGCwCGcK5agpE4rkvII+Zy3ql8Uf2XsoC1zq9z3su7"
    "HEOFWoSRIBqzkB5qEYHwPodXvqjeAqhghlpdQAxOwJZNRpy4PiHVv0ZHv7DGm0npBMD4WmkzDK804Q61hMBISQA4irFqCUZi"
    "uNXxU4CqUq2O1xwtb04JDdQKA2IN/dmsFmEkgO4soLlaRECU0I7yypc1WwDbEzIZCKAT/6GWYCSCXyWm+sOMqtWfDP3kN9QK"
    "A+RmDlFLMJznZC5TSwiQN6u/rB0AXlQrDJBG/E4twXCcIu5O1LSyeut3G/ZQkSBLUvQ2oud78hIcpG2hYf1JnimXGaStS8AO"
    "roaK/uyQl+AgbVrNBGZ6Vp6kTgC0Z7JaguEohfwpAfv/VaVW3U5+AIALuEItwXCS73OiWkLAvFDzD5mGN5qy0fmjj6uznsMS"
    "sOexES0DeTdh9WANXaixQiZTC2Anb6mVBkx7HqJILcJwiiY8kLDqDy9Sa4FcYR0XJo1T+E+1BMMpJiXwoLmc6/Wx8tHK4K2M"
    "M9X+N5zha/LyGoYdVDuhmac4FLGGduo8CJx1DGaFWoThAEcyPWGj/wALGFj7j5m7AGU8qVYbAh2YkrhenRE8bZmawOpP7TkA"
    "QJ1r5v+tVhsKx3Ov7RJgZKURj9JXLSIU/pXPxS3YKe+xhGM/V+eDEWMKuE9eQsOxtZmfg9X1e7iNl9R5ERI/5OtqCUZs+SnX"
    "qCWExBOUZfpz3Q3iZHYCACba8wAjI9fxI7WE0Hgi85/rXujYlRWJWgZZlR2cx8tqEUbMGMODidkNqya76MDWTG/U3QJYySy1"
    "6tBoxhOMUIswYsUlCa7+8ELm6k/WMfHkdgKgGU9yklqEERsu4qEEV/86HgHWxxHykctwbRND1flixIIL2CUvjWFaOd28OWaB"
    "XHq4to2z1GXPkDOW3fKSGK5l2ecz+7SYh9V5EzLN+RdXqkUYUm7hb7lsk+U0D3n9YF957ArfyrlVnT+GiALukpe/8K3Uz6Z4"
    "s+Xyo7Bf2m4BKaQZj8hLXhT2TDYn1Dcz3nPjwSlu5UnaqEUYkdKDV1OyY/SD2d6sb6pPV5anZPnMYi7kfbUIIyJGMIWOahGR"
    "8AWdsx2QV1/lXpmok4KycTBvcbFahBEBBdzKiymp/vBE9vMx6/91T0cnAKAlj3Ffgk6BMzLRkX+nasznwexv1z/bvy2raKRO"
    "RYR8xFW8pxZhhMTp3EsXtYgI2UJndma7oP4WwEaeU6ciUgYynZsTuwwqzTTlbp5NVfWHx7JX/9y4SP4gI3p7nQHqvDMC5SQ+"
    "kpeq6G14EK5rwEp5QqK3nYxPVdcnyRQzmXJ5iYrePqq/JZvLI7493KvOQQFNuJN3GKaWYfikgKtYwLhUdur+jwr/XwLQJ5Xx"
    "s4IKyplCL3U+Gp4ZwmvyMqSyXXQIzpEvyZOjs+1MoIW6JBt505XJlMlLj84CXcp3lTw5WlvKtYneMCJptOGnbJOXGq2NCtKh"
    "jVknT5DaPmVciiaQuEtLbmOTvLSobUluU/hznee/iwfU+SrnICYzn8tTsjbCTdownmVMoFgtRM6fKQ/2Cw+Tx7S42EJuopk6"
    "f41a9GEiW+WlIx5WStfcnJbPw5FXItlJdw8zmM88PmQdG9hNEw7iIHrTm2M5NDa/vhuYxO9Zo5Zh7ON4vstF1kE7wFTGBP+l"
    "F4Yetz7jm7TNoqAdF/Dr2OxU+AUPcLI6p1NPS25glrwsxM1OCMPVRSwOVfQjOT9sG8pE1sidvNc+5BbbTETEUO6xRn8Geycs"
    "h387RNEL8jy6uwGXMFPu6r22k0e5wCYOR0hvfsQH8nyPq10VlttbsTk00f/rSdEZMZqitIE/clJsRimSSidu4o3UzkzNxVaE"
    "ucvxr0OTPdGzplOZL3d6pa3jPkbTWFA1kk5vbuF5SuU5HHe7I8xM6MOekGR/4GO5RkO+F2LbxIuV8ABXBTkXO8U0ZhR38aE8"
    "T92w7bQLNzv+GZr0/E5mP5RR9K/yuhvPy51f08qYyU850SYRe6KAQdzCU2yX56NLNjnsbBkRmvSVeezW8pd9n3mHo6sUmG+z"
    "U54BmWwrz/JDTrBhwpwo4mhu4THWyvPNPStnUH7O9tLsfie0QzXfZQTbcriuA2sP/Hs7JzHnwKvD+Qd9Q1Lnnx28w3RmMIPV"
    "aikxpBPHMoxjGUYrtRRneZLz8vuAlwBwPv8KLQHPcHEOu5g14PMq2zpXT3RbHub00PQFxVJmMJf5zOcztRQpDTiEwziCQzna"
    "9l0IgOFMz+8DXgJAAXM4MrQkvMX5bKj3qrFVdilawMBq7zXgV9wSmr6g2cJ85jOP+czPvoN7gujFYRzG4RzGQOsWBciL+S8B"
    "9jbyfhmPhJiMBZydw+/iRfyGnkAp12Y4u+AHTHBwE6ilfMRilrCExSxhu1pOYDTmIHrThz77/mtN/HA4hVfy/Yi3SlLI+zV+"
    "dYNlHWOzH2m4T8XRtOH9OvrT1/FHxxeHrGExS1jMa7ykluKRhtzFYHrTzaZHRcCbnBjdza4JfTTzf3w/PLsqtDkLUdoSp/vG"
    "R8VmzUby7awoM7YBi0JP0Ov09Knyq85PGV1IjyizNQQOZb3ci2mwmd6yx2sjuZwdnB9y0enJjRQwPY+dTbpyJj1ZeuATc9jA"
    "OSGrDJMVnMQKtQifrOMVrrCp0aHzDRZEe8NGLI0kss3jqBwVDdo3HXglo6v89Rfy2OzVdjAk2iwNjdGp3p03CpunGPK+MaLE"
    "7eRrOem568Anyvjygb8W8og8e7xYOZdHn6Whcbvcn8m2SxSZ2pBPIkvgX2har55JVa7fxsEH/t6MefIMyt8mKLI0NAqYJvdo"
    "cm2G6pH35ZEmsr5VTldXu/4XVd7pR4k8k/Kz9xI3QaY9n8u9mlQ7TZWpBbwdYTLfr2en0+bVtix7s9p7Fzj1PGB3Ynr/VTlb"
    "7tdkWv0zZkLk9EiTupDeWdUMr3IaTE23/EaeUbnbj5RZGiIPyD2bPCtnsDZTX4g0uSsYkFXN0ftOgS+ttSqqsTN7yH2U2P0D"
    "OtiJPYHbQ/6zxR/HRNy4XsFBWfUUcDJjqm0Vsp9jHZkZGPb8CiXfk3s3WbY7DovfH4040R/TyaPSX8kzrH57TZ2dodKIhXIP"
    "J8kmqTMUoH/kGzXOyOGhYCZasEKeZdmtnGPV2RkyYa8iSZNto7Pf7AhildbH4e9DVoNh/NXTk89tfDdipfnyBDPUEkLmQRaq"
    "JSSGX8ZlZ6m2ggUfd3rU+pw8bmezU9RZGQFfkXs5Gfapx3ZwKNwUefLLcl7+WEzLKq/6xnTj0AoqmOvgJib505jVck8nwSTT"
    "f+uiiPcid8C6HJbKNuJ+yinn2Sq7pd4pz7q6bKw6GyPiJ3JPu28vqjOxJicLnPBmvU/Mv3Xg2i8O7JfWMKanya5K3PTfuuhq"
    "5/v4tFIODyYrgtuq6bVQ9wnMzHB+WM8VlbOkG/PPfduYlXJNDjsPR89D7FZLiIiVvKCW4Dh/YL5aQm16CM5w2V3lYJBMPFTt"
    "6gcP/P278hhe25L+ALAqV8q97bKto606AzPznwJnzM+628wPq12758CKwgIel2djdVuSigHA/TRli9zj7tqN6uyrO1uXCNzx"
    "0yyKurCj2rUnHXinOGZz0v5HnXkRM0XucVdtTpC7XQe7XfNOvh5xMQL4PofU+d4qflvtdesD/yrhUnYI1NZF9CMoWsI7XSrZ"
    "lDGOMrWIbNwviInPZdFTxD8OXPdFlQPFAM6PzfKglanqAAAUs1vudRft1+qMq492klNdL8uiqCHfYzMV7MrQd7pBnqF77QF1"
    "tgl4U+5192wpLdTZVj+K5R7L65kW2YQj6xg5nSDP1AoquE6daQJ+Kve6e3a2OtNy41mBa27zrDYOIeBgz+rd5VS5110zZ9qJ"
    "fQQzAjZ5fjZawN3ijF2qzjAJTW0UIC9bX2MEKxDCOT5zE7s5I+TiU5MmNMg6GJiNZyhkRMR6q/IE/xDeXcUeLqSLWoRDfKPG"
    "RrexpgGzI4+QO32do3eD8ImA9+6L2/xR/qvqjr0QznOisI5t3sNXI5/Z3oRb87i6ASO5pMoh55MZzcaIFe8nhvO6I8HjgZYp"
    "ZCvXU6EWkS+3RR4lt+fcS+rI/H2fmcXJB/7ak3ck0d31E4C9Mkz+u+qKfVWdVV4o4vXIHfXzHLX9vspn9vCNA39vyt2RHyFS"
    "krpJQPtp5dRxLTr7tzqjvNIn8iUfm6pM9s3GzGqfKueKKu+dFtG5x/vtDXU2CYn7Jq1xsLWed8HOgbDGAPayhO+E+v21KeaG"
    "nK7bXO1VAZNoc+DViwxiQmQjGMv5Q1TOiSEfqwU4wA2sUUvww78jjpef5vRos/b4RM1Q1Y+pITdQy3mFMYk9BSg3/iz/fY27"
    "/U2dRX7pFPnagNE5qGrDuhqfeizDVUfxOGWhaFzOz+Jwpouc8fIKFm/7LMcubay5MGKnPZ2TqtE1nvs/Xsd1/ZjI5gDVreBu"
    "RoQ0Acs9rpNXsThbGSPVGRQM/xex23L7bb282gbh12e5silX8AS7fKkq5S1+zHGpHfHPzFnyShZn+4U6e4KiKfMiddyvctTV"
    "n6nsooJSfpvDcGhLLub/9p0+nLtt5mV+yYVJaMqFwBB5JYuvvRHF+FBUv0cDmBnhSuYNdOeLHK9tRhfW13gqkJ32HEM/DqEf"
    "/eiVoTG/hxV8yqcsYTHvshDnZnBFSB8WqyXElPUMZnn4t4muQXo1f4/sXvClwJZONqknlBTTmOa0oCFlbGELO2K10VjcacsG"
    "tYRYUsF5PKUWETRRjgTMCSC0DeBhNlPBTl7nmwGew9aaPmFO7XCKopCesrhuidwiNtqRgHN8qh1QY+x/Kcf79kAzbuPDfd/3"
    "GXek5iSgbGyTV7b4WSS9fwUD2BqZE2f7nOU4qdY3buUon6lfXOMb/2XPBNgkr25xs3V0V2dKeFweoSO/4kvpwxm+8bU6r27G"
    "pfw3k/kZF9fRWWjOpxm+8WTSzhp5hYuXlXO+OkvCZVJkrlzp67nDVzN+Z/OM136FDVWuWc2YDNdknvJypTo75CyXV7l42U/U"
    "GRI2DXktMmf+zofOIl6p9X3bM/bNxtW6rjzDdORMk15301OdHXIUZ0nF1/4V8vK8WNAuskwv97UzYasqR4rstUxnERdkbMTW"
    "XuJ7YoalRT9QZ0UM+FRe6eJjH6VluthRke0a/CktfSk9namspYJSplfbM6CS5hnv+0mGK2+jtMoVG3yOUSSFz+TVLi62kX7q"
    "zIiOKyPbC+a+ANRmfyiTaRuxiRmv7M9PmMaz/JVrXTjhJRIsAOy1PZypzopoie44jptDTsmhrKpxx7d8tjvSRLS7L8XX8tnQ"
    "NhEUMi0i15aGvqiyE79mJRVUUM5cbqah2rkOsUxe9eJgUU6Tjw3FfByRe9fQK4L0dKBPHY8JjbqxAFDBrAAnmjtFn8imgSyk"
    "qzqxRkZsW9AVym3htU8dl3Ae2yO5U19esgU4sSTtk6G3cE4Uy37rQj3tYCZXUhbJnfrznOfjQw0jHEq5hHlqEWpuiayx9Q7t"
    "1Ik1avC5vAmus3KuVrs/Hvy/yFy+yPbijRmrIsv7+NkdaufHhUIejczpKzlWnVyjCukNAJPVro8TTXgjMsd/wdfUyTUOkNbl"
    "wE8lddMPr7Tjgwjd/yd7Yh8T0hkA3rap4LXpVmvHnDBtAUPUCTYg8nOj4mBz7XlUZnpGujSklInWDpCzLsIcj4ctpIva6fFl"
    "YMS/CAs5S53klJO2ALDUNoHJzhFsjDhLnudIdaJTzIaIc1traxmgdnj8GR75VtF7uJf+6mSnlDQFgA0coXa3G4yqdmxnNFbG"
    "wxytTngKibq9p7MtNgMld873eRavV3uLa2iiTnyq2CSvmNHYdk5Ru9otzuMLUVZt4I+MlC+TSgvpCABbGaF2tHucJegIVNrn"
    "/JZzQ35M2IRhXKB2s5gSeeUM3zZzgtrNdRHv1dij+BfNpAp28TovMZ1ZbAvsO1sziMEMYTCH0oDZHCNNoZrNtFJLCJkSzmKG"
    "WkRdxDsAwEieiMVknTI+4B0+5kMW8BnleX66JT3oxQD6058BdK723hwGqxMnZUvCN1DdyJnMUouom7gHADiRp2JWRHaxnNV8"
    "zirWsokv2M7maiGhmFa0pjWtaEc3utIj67zvuT6PHHWdrYmeFb+e03lPLSIb8Q8AMCTRe/nMS/k0pG2xaOGFw1pGMV8tIjsu"
    "jHXP5izWq0WEhgs5YHhhBSPjXv1dKX4zOY7FahEh4UIbLEzcKIH5s4AT+Egton5ccf9iTmKuWkQouJIDYZHMADiTk1mmFpEL"
    "7hS/VYzkdbWIEHAnB8IhiQHgBU5jnVpEbrhU/Eo4nalqEYHjUg6EQfICwAOcw1a1iFxxq/jt4gruUYsIGLdywKiPuxlLqVpE"
    "7rhW/Mq4kR+rRQSKazkQNElqAVTwY27Je5qYkTc3sFs+vzso+0ztTDHJycmdXKl2Zv64+fszmdMSMzMgSb+AXnCzBNZmPWfw"
    "kFpE/rjq/tc5ngVqEYHgag4ERTIC4HyOcfMZlbvFbxHDeUktIgDczYFgSEIAeJoTWaoW4Q2Xi98mzuT3ahG+cTkHDIC7OY8t"
    "ahHpZRyl8uEfP7Za7UAxav/7s1JuUjvQOMfpbaUcmTEWGmr/+7HVnKx2n1+S0AB9iiOZqRbhmST0gdOZ+lkcy2tqEX5JQgCA"
    "ZZzMn9QiPJKMHPCKuwHgHk5wdeAvqYxlh7xRmL+VqN0mpUjufy+2gy+rHWdk4miWyAtHvpbu8eMGcv/nb59wuNptwZGsBugc"
    "hvKMWkSeJCsH8sW9LsA/OCb++/ykmUL+iz3yX4ncbYfaYVIayf2fj+3iew6GrBRyLIvkhSVX26l2lpQmcv/nbguSuIF7Mhug"
    "MxjM39UiciSZOZAr7vye3s8xvKsWYeTDGCcmCDm0eUQINJf7PxcrcXGhrwG9eF1eeOqzPWonSWkh93/99hLd1W4yvNKA8TEf"
    "Ekz3/jGt5P7PbqWMp0jtJMMfJ/KJvCBlszSPArSWez+bzU/5wa2JoSkTYtwOSPMvTBu59+uyUibQWO0eIziO4wN5ocpsDdWu"
    "EdJW7v3MNo8hatcYQdOQ29glL1q1rZHaMULay71f23YzIdV5kmiOYLa8gNW0JmqnCOkg935Ne4+j1U4xwqQR42PWDmimdomQ"
    "jnLvV7Ud3E4DtUuM8OnL0/LCVmkt1O4Q0lnu/UqbRm+1O4zoGM1SeZHba63UrhDSRe79vbacsWpXGFHTLCadgdZqRwjpJvd+"
    "BbuZmOpWWKo5hOfkBbCN2glCusu9/xID1U4wlBRwNSulRbC92gVCekg9v4xL1Q4w4kBzbmOLrBh2UCdfSC+Z17cxnqbq5Bvx"
    "oSuTRdOFO6mTLqS3xOOlTE611406GMgUQXHsok62kD4Cfz/PYepkG/HlNN6NuEB2VSdZSN+IfT2TkeokG3GnkC/zWYSFsoc6"
    "wUL6RejnRVzp0BZkhpSGjI1sY9Ge6sQK6R+Rjz9jnE3yNfKjIWNZHEHhPEidUCEDrPIbcaYxN7Es5ALaR51IIQND9u0Svpbq"
    "/RaMAGjEOJaHWEhtIlA4tpRbbE8fIxgacz0fhVJM96R6T8CwDgaZy1j75TeCpZDRvBp4UV2mTpaYDYF79HnOtNF+IyyO5j5K"
    "Ayyu96kTJObxAH1ZxjSGqRNkJJ+DmEBJQIX2GnVixHwjID9uYWKqH6gaEdOWWwM4b2BbqhcDA3Rjt28vfsgtqd5VwRBRwCim"
    "+irA/6tOQgy414f/dvEQI9QJMNJNF/7D4/ZiK1P/+w/Qnc2evLeY2+moFm8YAEWM5gnK8irAuzlFLTsmXEZ5Xp7bw+OclerH"
    "p0Ys6c4PmJdz43WMWm6MuDnnEDCH76Z6AbURe47kV3xeTzFexPFqmTHjAtbW47PlTLCV/IYbFHE699axzdgyvm+bUWWgLb9g"
    "Y0aPbeavnGpN/jCwGVNh0ozzOYp+dKYFhWxmBe/zEm9TrhYWW5owkpEMoDstKGULK/mEOTzJTrUwwzAMwzAMwzCMBPD/ASb3"
    "i7+t2nCZAAAAAElFTkSuQmCC")
info_ico_data = (
    "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"
    "VORK5CYII=")

# Tkinter Icons
icon = tk.PhotoImage(data=b64decode(icon_data))
github_icon = tk.PhotoImage(data=b64decode(github_icon_data))
info_icon = tk.PhotoImage(data=b64decode(info_ico_data))

right_click_menu = tk.Menu(gui, tearoff=0)


# Modify the thread class to propagate exceptions
class PropagatingThread(threading.Thread):
    def run(self):
        self.exc = None
        try:
            if hasattr(self, '_Thread__target'):
                # Thread uses name mangling prior to Python 3.
                self.ret = self._Thread__target(*self._Thread__args, **self._Thread__kwargs)  # type: ignore
            else:
                self.ret = self._target(*self._args, **self._kwargs)  # type: ignore
        except BaseException as e:
            self.exc = e

    def join(self, timeout=None):
        super(PropagatingThread, self).join(timeout)
        if self.exc:
            raise self.exc
        return self.ret


class VerticalScrolledFrame(ttk.Frame):
    def __init__(self, parent, *args, **kw):
        ttk.Frame.__init__(self, parent, *args, **kw)

        # Create a canvas object and a vertical scrollbar for scrolling it.
        vscrollbar = ttk.Scrollbar(self, orient=VERTICAL)
        vscrollbar.pack(fill="y", side=RIGHT, expand=FALSE)
        self.canvas = tk.Canvas(self, bd=0, highlightthickness=0,
                                width=200, height=300,
                                yscrollcommand=vscrollbar.set)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=TRUE)
        vscrollbar.config(command=self.canvas.yview)

        # Reset the view
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

        # Create a frame inside the canvas which will be scrolled with it.
        self.interior = ttk.Frame(self.canvas)
        self.interior.bind('<Configure>', self._configure_interior)
        self.canvas.bind('<Configure>', self._configure_canvas)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.interior_id = self.canvas.create_window(0, 0, window=self.interior, anchor=NW)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _configure_interior(self, event):
        # Update the scrollbars to match the size of the inner frame.
        size = (self.interior.winfo_reqwidth(), self.interior.winfo_reqheight())
        self.canvas.config(scrollregion=(0, 0, size[0], size[1]))
        if self.interior.winfo_reqwidth() != self.canvas.winfo_width():
            # Update the canvas's width to fit the inner frame.
            self.canvas.config(width=self.interior.winfo_reqwidth())

    def _configure_canvas(self, event):
        if self.interior.winfo_reqwidth() != self.canvas.winfo_width():
            # Update the inner frame's width to fill the canvas.
            self.canvas.itemconfigure(self.interior_id, width=self.canvas.winfo_width())


def info_window():
    global gui, info_icon, TITLE, GITHUB, SHORT_NAME

    def close_info_window():
        info_window_gui.destroy()
        gui.attributes('-disabled', False)
        gui.update()
        gui.focus_force()

    objective_text = f"""{SHORT_NAME} assists you in mapping IPs from network logs to descriptive object names or domain names, making 
log analysis more straightforward."""
    features_text = f"""➞ Input Options:
      • Accepts Excel/CSV/Text files.
      • Direct input of single IP address, subnet, range, or list (Comma Separated).
➞ File Input Specifications:
      • Input Excel files could have multiple sheets, each must contain a column named 'Subnet'.
      • Reference Excel files require three columns: 'Tenant', 'Address object', and 'Subnet'.
➞ Search Methods:
      • Reference File: Searches for matches in a user-provided reference file.
      • Palo Alto: Connects via SSH to Panorama to fetch address objects.
      • Fortinet: Connects via REST API to FortiGate to retrieve address objects.
      • Cisco ACI: Connects via SSH to APIC to retrieve address objects based on a specified Class.
      • DNS: Resolves IPs to domain names using the system dns servers or a user-provided DNS servers.
➞ Palo Alto Panorama/Firewall Specifications:
      • Ensure Panorama/Firewall is reachable and has CLI access.
      • Leave "Vsys" field empty if you want to retrieve address objects from all virtual systems.
➞ Fortinet FortiGate Specifications:
      • Ensure FortiGate is reachable and you have REST API access.
      • Leave "Vdom" field empty if you want to retrieve address objects from all virtual domains.
➞ Cisco ACI Specifications:
      • Ensure APIC is reachable and has CLI access.
      • Specify the Class of the address objects to be searched.
      • The program searches the "dn" attribute exclusively.
➞ DNS Resolver:
      • Resolves IPs to domain names using the system dns servers or a user-provided DNS servers.
      • You can specify up to four DNS servers.
➞ Credentials Storage:
      • An option to save your credentials for future use.
      • Stored information are encrypted for security purposes.
➞ Output:
      • Results are exported to a new Excel (.xlsx) file for ease of access and analysis."""
    usage_text = f"""• Ensure the chosen search methods are accessible and correctly configured.
• Provide necessary credentials and remember to save them if needed.
• For Cisco ACI, you must specify the correct Class for targeted searches.
• Review the generated Excel file for mapped IPs based on the selected search methods."""

    info_window_gui = tk.Toplevel(gui)
    info_window_gui.title('Help')
    info_window_gui.focus_force()
    info_window_gui.grab_set()
    info_window_gui.focus_set()
    info_window_gui.transient(gui)
    gui.attributes('-disabled', True)

    info_window_gui.geometry(f'{665}x{580}+{gui.winfo_rootx() - 65}+{gui.winfo_rooty() - 100}')
    info_window_gui.resizable(False, False)
    info_window_gui.iconphoto(False, info_icon)

    info_frame = VerticalScrolledFrame(info_window_gui)
    info_frame.pack(expand=True, fill=tk.BOTH)

    # Objective
    objective_label1 = tk.Label(info_frame.interior, text="Objective:", font=("Arial bold", 10))
    objective_label1.pack(pady=(3, 1), padx=10, anchor="w")
    objective_label2 = ttk.Label(info_frame.interior, text=objective_text, font=("Arial", 10), wraplength=645,
                                 justify="left")
    objective_label2.pack(pady=0, padx=20, anchor="w")

    # Features
    features_label1 = tk.Label(info_frame.interior, text="Features:", font=("Arial bold", 10))
    features_label1.pack(pady=(3, 1), padx=10, anchor="w")
    features_label2 = ttk.Label(info_frame.interior, text=features_text, font=("Arial", 10), wraplength=645,
                                justify="left")
    features_label2.pack(pady=0, padx=20, anchor="w")

    # Usage
    usage_label1 = tk.Label(info_frame.interior, text="Usage Guidelines:", font=("Arial bold", 10))
    usage_label1.pack(pady=(3, 1), padx=10, anchor="w")
    usage_label2 = ttk.Label(info_frame.interior, text=usage_text, font=("Arial", 10), wraplength=645, justify="left")
    usage_label2.pack(pady=0, padx=20, anchor="w")

    auther_label = tk.Label(info_frame.interior, text=TITLE, font=("Arial bold", 10))
    auther_label.pack(pady=(10, 0), padx=10, anchor="w")

    github_label = tk.Label(info_frame.interior, text="GitHub", font=("Arial bold", 10), cursor="hand2",
                            foreground="blue")
    github_label.pack(pady=(0, 5), padx=10, anchor="w")
    github_label.bind("<Button-1>", lambda e: webbrowser.open_new(GITHUB))

    s2 = ttk.Style()
    s2.configure('close.TButton', font=('Arial bold', 10))
    close_button = ttk.Button(info_frame.interior, text="Close", cursor="hand2", command=close_info_window,
                              style='close.TButton')
    close_button.pack(pady=(0, 3), padx=10, anchor="center", ipady=3, ipadx=75)

    info_window_gui.protocol("WM_DELETE_WINDOW", close_info_window)
    info_window_gui.bind('<Escape>', close_info_window)

    info_window_gui.mainloop()


def clear_variables():
    global inputs, input_invalids, inputs_no, refs, ref_invalids, ref_no, output_file, outputs

    inputs = {}
    input_invalids = []
    inputs_no = 0

    refs = {}
    ref_invalids = []
    ref_no = 0

    output_file = ""
    outputs = {}


def set_info1(text, foreground, cursor):
    global info_label_1
    info_label_1.config(text=str(text), foreground=str(foreground), cursor=str(cursor))


def set_info2(text, foreground, cursor):
    global info_label_2
    info_label_2.config(text=str(text), foreground=str(foreground), cursor=str(cursor))


def right_click_fn():
    global gui, right_click_menu

    # Create a right click functions
    def cut():
        widget = gui.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<<Cut>>")

    def copy():
        widget = gui.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<<Copy>>")

    def paste():
        widget = gui.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<<Paste>>")

    def delete():
        widget = gui.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<Delete>")

    def select_all():
        widget = gui.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<<SelectAll>>")

    def clear_all():
        widget = gui.focus_get()
        if isinstance(widget, tk.Entry):
            widget.delete(0, "end")

    # Create a right click menu
    right_click_menu = tk.Menu(gui, tearoff=0)
    right_click_menu.add_command(label="Cut", command=cut)
    right_click_menu.add_command(label="Copy", command=copy)
    right_click_menu.add_command(label="Paste", command=paste)
    right_click_menu.add_command(label="Delete", command=delete)
    right_click_menu.add_separator()
    right_click_menu.add_command(label="Select All", command=select_all)
    right_click_menu.add_command(label="Clear All", command=clear_all)


def right_click(event):
    global right_click_menu
    event.widget.focus()
    right_click_menu.tk_popup(event.x_root, event.y_root)


def disable_buttons():
    global button_panorama, button_forti, button_aci, button_dns, button_start
    button_panorama.config(state="disabled", cursor="arrow")
    button_forti.config(state="disabled", cursor="arrow")
    button_aci.config(state="disabled", cursor="arrow")
    button_dns.config(state="disabled", cursor="arrow")
    button_start.config(state="disabled", cursor="arrow")


def enable_buttons():
    global button_panorama, button_forti, button_aci, button_dns, button_start, flags

    if (button_panorama['text'] == "Connect" and
            button_aci['text'] == "Connect" and
            button_forti['text'] == "Connect" and
            button_dns['text'] == "Check"):
        button_panorama.config(state="normal", cursor="hand2")
        button_forti.config(state="normal", cursor="hand2")
        button_aci.config(state="normal", cursor="hand2")
        button_dns.config(state="normal", cursor="hand2")

        if True in flags:
            button_start.config(state="normal", cursor="hand2")


def disable_buttons_all(event=None):
    global gui, input_entry, input_button, ref_entry, ref_button, ref_checkbox, button_panorama, panorama_checkbox, \
        button_forti, forti_checkbox, button_aci, aci_checkbox, button_dns, dns_checkbox

    gui.protocol("WM_DELETE_WINDOW", stop_process)
    input_entry.config(state="disabled", cursor="arrow")
    input_button.config(state="disabled", cursor="arrow")
    ref_entry.config(state="disabled", cursor="arrow")
    ref_button.config(state="disabled", cursor="arrow")
    ref_checkbox.config(state="disabled", cursor="arrow")
    button_panorama.config(state="disabled", cursor="arrow")
    panorama_checkbox.config(state="disabled", cursor="arrow")
    button_forti.config(state="disabled", cursor="arrow")
    forti_checkbox.config(state="disabled", cursor="arrow")
    button_aci.config(state="disabled", cursor="arrow")
    aci_checkbox.config(state="disabled", cursor="arrow")
    button_dns.config(state="disabled", cursor="arrow")
    dns_checkbox.config(state="disabled", cursor="arrow")


def enable_buttons_all(event=None):
    global gui, input_entry, input_button, ref_entry, ref_button, ref_checkbox, button_panorama, panorama_checkbox, \
        button_forti, forti_checkbox, button_aci, aci_checkbox, button_dns, dns_checkbox

    gui.protocol("WM_DELETE_WINDOW", lambda: close_gui_window())
    input_entry.config(state="normal", cursor="xterm")
    input_button.config(state="normal", cursor="hand2")
    ref_entry.config(state="normal", cursor="xterm")
    ref_button.config(state="normal", cursor="hand2")
    ref_checkbox.config(state="normal", cursor="hand2")
    button_panorama.config(state="normal", cursor="hand2")
    panorama_checkbox.config(state="normal", cursor="hand2")
    button_forti.config(state="normal", cursor="hand2")
    forti_checkbox.config(state="normal", cursor="hand2")
    button_aci.config(state="normal", cursor="hand2")
    aci_checkbox.config(state="normal", cursor="hand2")
    button_dns.config(state="normal", cursor="hand2")
    dns_checkbox.config(state="normal", cursor="hand2")


def check_methods(event=None, *arg):
    global gui, flags, start_flag, button_start, button_panorama, button_aci, button_forti, button_dns, \
        ref_checkbox_var, panorama_checkbox_var, aci_checkbox_var, forti_checkbox_var, dns_checkbox_var

    if ref_checkbox_var.get():
        flags[0] = True
    else:
        flags[0] = False

    if panorama_checkbox_var.get():
        flags[1] = True
    else:
        flags[1] = False

    if aci_checkbox_var.get():
        flags[2] = True
    else:
        flags[2] = False

    if forti_checkbox_var.get():
        flags[3] = True
    else:
        flags[3] = False

    if dns_checkbox_var.get():
        flags[4] = True
    else:
        flags[4] = False

    if True in flags:
        if not start_flag:
            set_info1("INFO", "green", "arrow")
            set_info2("You must enter input file / IP and at least one valid searching method", "green", "arrow")
        if (str(button_panorama["state"]) == "normal" and
                str(button_aci["state"]) == "normal" and
                str(button_forti["state"]) == "normal" and
                str(button_dns["state"]) == "normal"):
            button_start.config(state="normal", cursor="hand2")
        gui.bind("<Return>", start)
        return True
    else:
        if not start_flag:
            set_info1("INFO", "green", "arrow")
            set_info2("Please select one searching method at least", "red", "arrow")
        button_start.config(state="disabled", cursor="arrow")
        gui.bind("<Return>", lambda e: False)
        return False


# Create a function to browse the input file
def browse_input():
    global input_file, input_entry

    # Open a file dialog
    filename = filedialog.askopenfilename(title="Select an Input File",
                                          filetypes=(("Supported Files", "*.xlsx;*.csv;*.txt;*.xls;*.xlsm;*.xlsb;*.odf;"
                                                                         "*.ods;*.odt"),
                                                     ("All Files", "*.*")))
    # Check if a file is selected
    if filename != "":
        # Insert the file path in the entry
        input_entry.delete(0, "end")
        input_entry.insert(0, filename)
        input_file = filename


# Check if the input file is valid
def check_input_file(*arg):
    global input_file, input_entry, inputs, inputs_no, input_invalids, logs, start_flag, separators, inputs_bkp

    if "input_entry" in str(input_entry.focus_get()):
        return False

    if input_file == input_entry.get().strip() and (not start_flag) and inputs_no:
        set_info1("INFO", "green", "arrow")
        set_info2(f"The input is valid.", "green", "arrow")
        return True

    input_file = input_entry.get().strip()
    input_type = get_subnet_type(input_file)

    if input_file == "":
        set_info1("ERROR", "red", "arrow")
        set_info2("Please select an input file.", "red", "arrow")
        if start_flag:
            logs.set("ERROR - Input file is not selected.")
        return False

    elif input_type != "Invalid":
        if input_type != "List":
            inputs = {"Result": [[input_file, input_type]]}
            inputs_no = 1
            input_invalids = []
        if any(sep in input_file for sep in separators):
            inputs = {}
            inputs_no = 0
            input_invalids = []

            sep = [sep for sep in separators if sep in input_file][0]
            input_file = input_file.split(sep)

            for i, subnet in enumerate(input_file):
                subnet = subnet.strip()
                subnet_type = get_subnet_type(subnet)
                if subnet_type == "Invalid":
                    input_invalids.append(["Input", subnet, i + 1])
                    continue
                inputs.setdefault("Result", []).append([subnet, subnet_type])
                inputs_no += 1

        set_info1("INFO", "green", "arrow")
        if input_type != "IP":
            tmp = input_type.lower()
        else:
            tmp = input_type
        set_info2(f"The input {tmp} is valid.", "green", "arrow")
        logs.set(f"INFO - Input {tmp} is valid ({input_file}).")

        inputs_bkp = copy.deepcopy(inputs)
        for sheet in inputs.keys():
            for row in inputs[sheet]:
                if isinstance(row[0], str):
                    row[0] = convert_address_object(row[0], row[1])
        return True

    elif input_type == "Invalid":
        if isinstance(input_file, str):
            excel_ext = (".xls", ".xlsx", ".xlsm", ".xlsb", ".odf", ".ods", ".odt")
            csv_ext = (".csv",)
            text_ext = (".txt",)

            data = {}
            inputs = {}
            input_invalids = []
            inputs_no = 0

            # Check if the input file exists
            if not os.path.isfile(input_file):
                set_info1("ERROR", "red", "arrow")
                set_info2("The input file does not exist.", "red", "arrow")
                if start_flag:
                    logs.set(f"ERROR - Input file does not exist ({input_file}).")
                return False

            # Check if the input file is an Excel file
            elif not (input_file.lower().endswith(excel_ext)
                      or input_file.lower().endswith(csv_ext)
                      or input_file.lower().endswith(text_ext)):
                set_info1("ERROR", "red", "arrow")
                set_info2("The input file extension is not a valid type.", "red", "arrow")
                if start_flag:
                    logs.set(f"ERROR - Input file extension is not a valid type ({input_file}).")
                return False

            # Check if the input file is empty
            elif os.stat(input_file).st_size == 0:
                set_info1("ERROR", "red", "arrow")
                set_info2("The input file is empty.", "red", "arrow")
                if start_flag:
                    logs.set(f"ERROR - Input file is empty ({input_file}).")
                return False

            # Check if the input file is an Excel file
            elif input_file.lower().endswith(excel_ext):
                sheet_names = []

                # Find the engine to read the Excel file and get the sheet names
                for engine in [None, "xlrd", "openpyxl", "odf", "pyxlsb"]:
                    try:
                        sheet_names = pd.ExcelFile(input_file, engine=engine).sheet_names
                        break
                    except:
                        pass

                if not sheet_names:
                    set_info1("ERROR", "red", "arrow")
                    set_info2("The input file has no sheets.", "red", "arrow")
                    if start_flag:
                        logs.set(f"ERROR - Input file has no sheets ({input_file}).")
                    return False

                # Get sheet data
                for sheet_name in sheet_names:
                    df = pd.read_excel(input_file, sheet_name=sheet_name, engine=engine)
                    # check if there is a column named "Subnet"
                    if "Subnet" in df.columns:
                        data[sheet_name] = df["Subnet"].tolist()
                    elif "subnet" in df.columns:
                        data[sheet_name] = df["subnet"].tolist()

                if not data:
                    set_info1("ERROR", "red", "arrow")
                    set_info2("The input file has no (Subnet) column.", "red", "arrow")
                    if start_flag:
                        logs.set(f"ERROR - Input file has no (Subnet) column ({input_file}).")
                    return False

                for sheet_name in data.keys():
                    for i, subnet in enumerate(data[sheet_name]):
                        subnet = str(subnet).strip()
                        subnet_type = get_subnet_type(subnet)

                        if subnet_type == "Invalid":
                            if subnet == "nan":
                                input_invalids.append([sheet_name, "Empty", i + 2])
                                subnet = ""
                            else:
                                input_invalids.append([sheet_name, subnet, i + 2])
                            # continue
                        else:
                            inputs_no += 1

                        inputs.setdefault(sheet_name, []).append([subnet, subnet_type])

                if inputs_no == 0:
                    set_info1("ERROR", "red", "arrow")
                    set_info2("The input file has no valid subnets.", "red", "arrow")
                    if start_flag:
                        logs.set(f"ERROR - Input file has no valid subnets ({input_file}).")
                    return False

                set_info1("INFO", "green", "arrow")
                set_info2(f"The input file is valid ({inputs_no} subnets).", "green", "arrow")
                logs.set(f"INFO - Input file is valid ({input_file}).")
                inputs_bkp = copy.deepcopy(inputs)
                for sheet in inputs.keys():
                    for row in inputs[sheet]:
                        if isinstance(row[0], str):
                            row[0] = convert_address_object(row[0], row[1])
                return True

            elif input_file.lower().endswith(csv_ext):
                try:
                    df = pd.read_csv(input_file)
                except:
                    set_info1("ERROR", "red", "arrow")
                    set_info2("The input file is not valid.", "red", "arrow")
                    if start_flag:
                        logs.set(f"ERROR - Input file is not valid ({input_file}).")
                    return False

                # Check if there is a column called "Subnet"
                if "Subnet" in df.columns:
                    subnets = df["Subnet"].tolist()
                elif "subnet" in df.columns:
                    subnets = df["subnet"].tolist()
                else:
                    set_info1("ERROR", "red", "arrow")
                    set_info2("The input file has no (Subnet) column.", "red", "arrow")
                    if start_flag:
                        logs.set(f"ERROR - Input file has no (Subnet) column ({input_file}).")
                    return False

                for i, subnet in enumerate(subnets):
                    subnet = str(subnet).strip()
                    subnet_type = get_subnet_type(subnet)

                    if subnet_type == "Invalid":
                        if subnet == "nan":
                            input_invalids.append(["CSV", "Empty", i + 2])
                            subnet = ""
                        else:
                            input_invalids.append(["CSV", subnet, i + 2])
                        # continue
                    else:
                        inputs_no += 1

                    inputs.setdefault("CSV", []).append([subnet, subnet_type])

                if inputs_no == 0:
                    set_info1("ERROR", "red", "arrow")
                    set_info2("The input file has no valid subnets.", "red", "arrow")
                    if start_flag:
                        logs.set(f"ERROR - Input file has no valid subnets ({input_file}).")
                    return False

                set_info1("INFO", "green", "arrow")
                set_info2(f"The input file is valid ({inputs_no} subnets).", "green", "arrow")
                logs.set(f"INFO - Input file is valid ({input_file}).")
                inputs_bkp = copy.deepcopy(inputs)
                for sheet in inputs.keys():
                    for row in inputs[sheet]:
                        if isinstance(row[0], str):
                            row[0] = convert_address_object(row[0], row[1])
                return True

            elif input_file.lower().endswith(text_ext):
                try:
                    with open(input_file, "r") as f:
                        data = f.readlines()
                except:
                    set_info1("ERROR", "red", "arrow")
                    set_info2("The input file is not valid.", "red", "arrow")
                    if start_flag:
                        logs.set(f"ERROR - Input file is not valid ({input_file}).")
                    return False

                for i, subnet in enumerate(data):
                    subnet = str(subnet).strip()
                    subnet_type = get_subnet_type(subnet)

                    if subnet_type == "Invalid":
                        if subnet == "nan":
                            input_invalids.append(["TXT", "Empty", i + 2])
                            subnet = ""
                        else:
                            input_invalids.append(["TXT", subnet, i + 2])
                        # continue
                    else:
                        inputs_no += 1

                    inputs.setdefault("TXT", []).append([subnet, subnet_type])

                if inputs_no == 0:
                    set_info1("ERROR", "red", "arrow")
                    set_info2("The input file has no valid subnets.", "red", "arrow")
                    if start_flag:
                        logs.set(f"ERROR - Input file has no valid subnets ({input_file}).")
                    return False

                set_info1("INFO", "green", "arrow")
                set_info2(f"The input file has ({inputs_no} subnets).", "green", "arrow")
                logs.set(f"INFO - Input file is valid ({input_file}).")
                inputs_bkp = copy.deepcopy(inputs)
                for sheet in inputs.keys():
                    for row in inputs[sheet]:
                        if isinstance(row[0], str):
                            row[0] = convert_address_object(row[0], row[1])
                return True

        else:
            set_info1("ERROR", "red", "arrow")
            set_info2("The input file is not valid.", "red", "arrow")
            if start_flag:
                logs.set(f"ERROR - Input file is not valid ({input_file}).")

    return False


# Create a function to browse the reference file
def browse_ref():
    global ref_file, ref_entry

    # Open a file dialog
    filename = filedialog.askopenfilename(title="Select a Reference File",
                                          filetypes=(("Supported Files", "*.xlsx;*.csv;*.xls;*.xlsm;*.xlsb;*.odf;"
                                                                         "*.ods;*.odt"),
                                                     ("All Files", "*.*")))
    # Check if a file is selected
    if filename != "":
        # Insert the file path in the entry
        ref_entry.delete(0, "end")
        ref_entry.insert(0, filename)
        ref_file = filename


# Check if the reference file is valid
def check_ref_file(*arg):
    global ref_file, ref_entry, refs, ref_no, ref_invalids, logs, start_flag, separators, refs_bkp

    if "ref_entry" in str(ref_entry.focus_get()):
        return False

    if ref_file == ref_entry.get().strip() and (not start_flag) and ref_no:
        set_info1("INFO", "green", "arrow")
        set_info2(f"The reference file is valid.", "green", "arrow")
        return True

    ref_file = ref_entry.get().strip()

    # Check if the reference file is selected
    if ref_file == "":
        if flags[0]:
            set_info1("ERROR", "red", "arrow")
            set_info2("Please select a reference file.", "red", "arrow")
            if start_flag:
                logs.set("ERROR - Reference file is not selected.")
        return False

    elif isinstance(ref_file, str):
        excel_ext = (".xls", ".xlsx", ".xlsm", ".xlsb", ".odf", ".ods", ".odt")
        csv_ext = (".csv",)
        # text_ext = (".txt",)

        column_1 = ("Tenant", "tenant", "Tenant / Vsys", "Zone", "zone", "vlan", "VLAN")
        column_2 = ("Address object", "address object", "Address Object", "address Object",
                    "Name", "name", "Object Name", "object name", "Object name", "object Name")
        column_3 = ("Subnet", "subnet", "IP", "ip", "IP Address", "ip address", "IP address",
                    "ip Address", "IP/Subnet", "ip/subnet", "IP/subnet", "ip/Subnet")

        data = {}
        refs = {}
        ref_invalids = []
        ref_no = 0

        # Check if the reference file exists
        if not os.path.isfile(ref_file):
            set_info1("ERROR", "red", "arrow")
            set_info2("The reference file does not exist.", "red", "arrow")
            if start_flag:
                logs.set(f"ERROR - Reference file does not exist ({ref_file}).")
            return False

        # Check if the reference file is an Excel file
        if not (ref_file.lower().endswith(excel_ext)
                or ref_file.lower().endswith(csv_ext)):
            set_info1("ERROR", "red", "arrow")
            set_info2("The reference file extension is not a valid type.", "red", "arrow")
            if start_flag:
                logs.set(f"ERROR - Reference file extension is not a valid type ({ref_file}).")
            return False

        # Check if the reference file is empty
        if os.stat(ref_file).st_size == 0:
            set_info1("ERROR", "red", "arrow")
            set_info2("The reference file is empty.", "red", "arrow")
            if start_flag:
                logs.set(f"ERROR - Reference file is empty ({ref_file}).")
            return False

        # Check if the reference file is an Excel file
        if ref_file.lower().endswith(excel_ext):
            sheet_names = []

            # Find the engine to read the Excel file and get the sheet names
            for engine in [None, "xlrd", "openpyxl", "odf", "pyxlsb"]:
                try:
                    sheet_names = pd.ExcelFile(ref_file, engine=engine).sheet_names
                    break
                except:
                    pass

            if not sheet_names:
                set_info1("ERROR", "red", "arrow")
                set_info2("The reference file has no sheets.", "red", "arrow")
                if start_flag:
                    logs.set(f"ERROR - Reference file has no sheets ({ref_file}).")
                return False

            # Get sheet data
            for sheet_name in sheet_names:
                df = pd.read_excel(ref_file, sheet_name=sheet_name, engine=engine)
                # Check if there are columns called "Tenant", "Address object", and "Subnet"
                if (any(col in df.columns for col in column_1)
                        and any(col in df.columns for col in column_2)
                        and any(col in df.columns for col in column_3)):
                    # Get the matching columns
                    _1 = [col for col in column_1 if col in df.columns][0]  # Tenant
                    _2 = [col for col in column_2 if col in df.columns][0]  # Address object
                    _3 = [col for col in column_3 if col in df.columns][0]  # Subnet
                    data[sheet_name] = df[[_1, _2, _3]].values.tolist()  # [Tenant, Address object, Subnet]

            if not data:
                set_info1("ERROR", "red", "arrow")
                set_info2("The reference file has no valid columns.", "red", "arrow")
                if start_flag:
                    logs.set(f"ERROR - Reference file has no valid columns ({ref_file}).")
                return False

            for sheet_name in data.keys():
                for i, row in enumerate(data[sheet_name]):
                    ip = str(row[2]).strip()
                    subnet_type = get_subnet_type(ip)

                    if subnet_type == "Invalid":
                        if ip == "nan":
                            ref_invalids.append([sheet_name, "Empty", i + 2])
                            row[2] = ""
                        else:
                            ref_invalids.append([sheet_name, ip, i + 2])
                        # continue
                    else:
                        ref_no += 1

                    row.append(subnet_type)
                    row = [str(cell).strip() for cell in row]

                    if row[0] == "nan":
                        row[0] = ""
                    if row[1] == "nan":
                        row[1] = ""

                    refs.setdefault(sheet_name, []).append(row)

            if ref_no == 0:
                set_info1("ERROR", "red", "arrow")
                set_info2("The reference file has no valid subnets.", "red", "arrow")
                if start_flag:
                    logs.set(f"ERROR - Reference file has no valid subnets ({ref_file}).")
                return False

            ref_checkbox_var.set(1)
            set_info1("INFO", "green", "arrow")
            set_info2(f"The reference file is valid ({ref_no} subnets).", "green", "arrow")
            logs.set(f"INFO - Reference file is valid ({ref_file}).")
            refs_bkp = copy.deepcopy(refs)
            for sheet in refs.keys():
                for row in refs[sheet]:
                    if isinstance(row[0], str):
                        row[0] = convert_address_object(row[0], row[1])
            return True

        elif ref_file.lower().endswith(csv_ext):
            try:
                df = pd.read_csv(ref_file)
            except:
                set_info1("ERROR", "red", "arrow")
                set_info2("The reference file is not valid.", "red", "arrow")
                if start_flag:
                    logs.set(f"ERROR - Reference file is not valid ({ref_file}).")
                return False

            # Check if there are columns called "Tenant", "Address object", and "Subnet"
            if (any(col in df.columns for col in column_1)
                    and any(col in df.columns for col in column_2)
                    and any(col in df.columns for col in column_3)):
                # Get the matching columns
                _1 = [col for col in column_1 if col in df.columns][0]  # Tenant
                _2 = [col for col in column_2 if col in df.columns][0]  # Address object
                _3 = [col for col in column_3 if col in df.columns][0]  # Subnet
                data = df[[_1, _2, _3]].values.tolist()  # [Tenant, Address object, Subnet]
            else:
                set_info1("ERROR", "red", "arrow")
                set_info2("The reference file has no valid columns.", "red", "arrow")
                if start_flag:
                    logs.set(f"ERROR - Reference file has no valid columns ({ref_file}).")
                return False

            for i, row in enumerate(data):
                ip = str(row[2]).strip()
                subnet_type = get_subnet_type(ip)

                if subnet_type == "Invalid":
                    if ip == "nan":
                        ref_invalids.append(["CSV", "Empty", i + 2])
                        row[2] = ""
                    else:
                        ref_invalids.append(["CSV", ip, i + 2])
                    # continue
                else:
                    ref_no += 1

                row.append(subnet_type)
                row = [str(cell).strip() for cell in row]

                if row[0] == "nan":
                    row[0] = ""
                if row[1] == "nan":
                    row[1] = ""

                refs.setdefault("CSV", []).append(row)

            ref_checkbox_var.set(1)
            set_info1("INFO", "green", "arrow")
            set_info2(f"The reference file is valid ({ref_no} subnets).", "green", "arrow")
            logs.set(f"INFO - Reference file is valid ({ref_file}).")
            refs_bkp = copy.deepcopy(refs)
            for sheet in refs.keys():
                for row in refs[sheet]:
                    if isinstance(row[0], str):
                        row[0] = convert_address_object(row[0], row[1])
            return True

    else:
        set_info1("ERROR", "red", "arrow")
        set_info2("The reference file is not valid.", "red", "arrow")
        if start_flag:
            logs.set(f"ERROR - Reference file is not valid ({ref_file}).")

    return False


def import_credentials(app=""):
    global KEY, panorama_ip, panorama_username, panorama_password, panorama_vsys, aci_ip, aci_username, aci_password, \
        aci_class, forti_ip, forti_port, forti_username, forti_password, forti_vdom, logs

    output = [False for _ in range(4)]
    # Get the Panorama credentials
    cipher_suite = Fernet(KEY)

    while True:
        if app not in ["panorama", ""]:
            break

        # Check if the Panorama credentials file exists
        if os.path.isfile('panorama.cfg'):
            try:
                # Read the Panorama credentials file
                with open('panorama.cfg', 'r') as f:
                    encrypted_credentials = f.read()
                plain_text = cipher_suite.decrypt(encrypted_credentials.encode('utf-8'))
            except:
                # Delete the Panorama credentials file
                # os.remove('panorama.cfg')
                # logs.set("ERROR - Panorama credentials file (panorama.cfg) is corrupted and has been deleted.")

                os.rename('panorama.cfg', f'{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}_panorama.cfg.bkp')
                logs.set("ERROR - Panorama credentials file (panorama.cfg) is corrupted and has been renamed.")

                output[0] = False
                break

            plain_text = plain_text.decode('utf-8')
            plain_text = plain_text.split('\n')

            if len(plain_text) == 5 and plain_text[0] == "PAN":
                # Set the Panorama credentials
                panorama_ip = plain_text[1]
                panorama_username = plain_text[2]
                panorama_password = plain_text[3]
                panorama_vsys = plain_text[4]
                logs.set("INFO - Panorama credentials are imported from a file.")
                output[0] = True
            else:
                os.remove('panorama.cfg')
                logs.set("ERROR - Panorama credentials file (panorama.cfg) is corrupted and has been deleted.")
                output[0] = False
                break
        break

    while True:
        if app not in ["aci", ""]:
            break

        # Check if the ACI credentials file exists
        if os.path.isfile('aci.cfg'):
            try:
                # Read the ACI credentials file
                with open('aci.cfg', 'r') as f:
                    encrypted_credentials = f.read()
                plain_text = cipher_suite.decrypt(encrypted_credentials.encode('utf-8'))
            except:
                # Delete the ACI credentials file
                # os.remove('aci.cfg')
                # logs.set("ERROR - ACI credentials file (aci.cfg) is corrupted and has been deleted.")

                os.rename('aci.cfg', f'{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}_aci.cfg.bkp')
                logs.set("ERROR - ACI credentials file (aci.cfg) is corrupted and has been renamed.")

                output[1] = False
                break

            plain_text = plain_text.decode('utf-8')
            plain_text = plain_text.split('\n')

            if len(plain_text) == 5 and plain_text[0] == "ACI":
                # Set the ACI credentials
                aci_ip = plain_text[1]
                aci_username = plain_text[2]
                aci_password = plain_text[3]
                aci_class = plain_text[4]
                logs.set("INFO - ACI credentials are imported from a file.")
                output[1] = True
            else:
                os.remove('aci.cfg')
                logs.set("ERROR - ACI credentials file (aci.cfg) is corrupted and has been deleted.")
                output[1] = False
                break
        break

    while True:
        if app not in ["forti", ""]:
            break

        # Check if the Forti credentials file exists
        if os.path.isfile('fortigate.cfg'):
            try:
                # Read the Forti credentials file
                with open('fortigate.cfg', 'r') as f:
                    encrypted_credentials = f.read()
                plain_text = cipher_suite.decrypt(encrypted_credentials.encode('utf-8'))
            except:
                # Delete the Forti credentials file
                # os.remove('fortigate.cfg')
                # logs.set("ERROR - Forti credentials file (fortigate.cfg) is corrupted and has been deleted.")

                os.rename('fortigate.cfg', f'{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}_fortigate.cfg.bkp')
                logs.set("ERROR - Forti credentials file (fortigate.cfg) is corrupted and has been renamed.")

                output[2] = False
                break

            plain_text = plain_text.decode('utf-8')
            plain_text = plain_text.split('\n')

            if len(plain_text) == 6 and plain_text[0] == "FORTI":
                # Set the Forti credentials
                forti_ip = plain_text[1]
                forti_port = plain_text[2]
                forti_username = plain_text[3]
                forti_password = plain_text[4]
                forti_vdom = plain_text[5]
                logs.set("INFO - Forti credentials are imported from a file.")
                output[2] = True
            else:
                os.remove('fortigate.cfg')
                logs.set("ERROR - Forti credentials file (fortigate.cfg) is corrupted and has been deleted.")
                output[2] = False
                break
        break

    while True:
        if app not in ["dns", ""]:
            break

        # Check if the DNS credentials file exists
        if os.path.isfile('dns.cfg'):
            try:
                # Read the DNS credentials file
                with open('dns.cfg', 'r') as f:
                    encrypted_credentials = f.read()
                plain_text = cipher_suite.decrypt(encrypted_credentials.encode('utf-8'))
            except:
                # Delete the DNS credentials file
                # os.remove('dns.cfg')
                # logs.set("ERROR - DNS credentials file (dns.cfg) is corrupted and has been deleted.")

                os.rename('dns.cfg', f'{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}_dns.cfg.bkp')
                logs.set("ERROR - DNS credentials file (dns.cfg) is corrupted and has been renamed.")

                output[3] = False
                break

            plain_text = plain_text.decode('utf-8')
            plain_text = plain_text.split('\n')

            if len(plain_text) == 5 and plain_text[0] == "DNS":
                # Set the DNS credentials
                dns_servers = plain_text[1:]
                logs.set("INFO - DNS credentials are imported from a file.")
                output[3] = True
            else:
                os.remove('dns.cfg')
                logs.set("ERROR - DNS credentials file (dns.cfg) is corrupted and has been deleted.")
                output[3] = False
                break
        break

    return output


################################################################

def connect_to_panorama(ip=panorama_ip,
                        username=panorama_username,
                        password=panorama_password,
                        vsys=panorama_vsys,
                        window=None):
    global gui, ssh_panorama, panorama_ip, panorama_username, panorama_password, panorama_vsys, panorama_addresses, logs

    disable_buttons()

    if ssh_panorama is not None:
        disconnect_from_panorama()

    # Check if the Panorama credentials are correct
    if ip.strip() == "" or username.strip() == "" or password == "":
        button_panorama.config(text="Connect")
        enable_buttons()
        disconnect_from_panorama(False)
        logs.set("ERROR - Palo Alto credentials are empty.")
        return False

    # Check if the Panorama credentials are correct
    ssh_panorama = paramiko.SSHClient()
    ssh_panorama.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    button_panorama.config(cursor="wait", text="Connecting")
    try:
        # Connect to Panorama
        ssh_panorama_thread = PropagatingThread(target=ssh_panorama.connect, args=(ip, 22, username, password),
                                                kwargs={'timeout': 10,
                                                        'allow_agent': False,
                                                        'look_for_keys': False},
                                                daemon=True,
                                                name="ssh_panorama_thread")
        ssh_panorama_thread.start()

        while ssh_panorama_thread.is_alive():
            gui.update()
            if window is not None:
                window.update()
            time.sleep(0.1)

        ssh_panorama_thread.join()

        # Check if the connection is still open
        if ssh_panorama.get_transport().is_active():
            try:
                transport = ssh_panorama.get_transport()
                transport.send_ignore()
            except Exception as e:
                logs.set(f"ERROR - Failed connecting to Palo Alto: {str(e)}")
                button_panorama.config(text="Connect")
                enable_buttons()
                disconnect_from_panorama()
                return False

        # Update the Panorama status label
        label_panorama_status.config(text="Connected", foreground="green")
        logs.set(f"INFO - Connected to Palo Alto Device ({ip}).")

        try:
            if isinstance(ssh_panorama, paramiko.SSHClient) and ssh_panorama.get_transport().is_active():
                export_panorama_thread = PropagatingThread(target=export_panorama_address_objects,
                                                           args=(ip, vsys),
                                                           daemon=True,
                                                           name="export_panorama_thread")
                export_panorama_thread.start()

                while export_panorama_thread.is_alive():
                    gui.update()
                    if window is not None:
                        window.update()
                    time.sleep(0.1)

                export_panorama_thread.join()

        except Exception as e:
            panorama_addresses = []
            logs.set(f"ERROR - Failed exporting address objects from Palo Alto: {str(e)}")

        button_panorama.config(text="Connect")
        enable_buttons()
        return True

    except Exception as e:
        logs.set(f"ERROR - Failed connecting to Palo Alto: {str(e)}")
        button_panorama.config(text="Connect")
        enable_buttons()
        disconnect_from_panorama(False)
        return False


def export_panorama_address_objects(ip=panorama_ip, vsys=panorama_vsys):
    global ssh_panorama, panorama_addresses, panorama_addresses_bkp, logs, panorama_checkbox_var, gui

    try:
        ssh_panorama_shell = None
        panorama_addresses = []
        vsys = re.sub(" ", "", vsys.strip())

        if ssh_panorama == None:
            return False

        elif isinstance(ssh_panorama, paramiko.SSHClient):
            # Execute the commands
            try:
                ssh_panorama_shell = ssh_panorama.invoke_shell()  # Invoke the shell on Panorama
            except Exception as e:
                logs.set(f"ERROR - Exception invoking shell on Palo Alto. {str(e)}")
                return False

        ssh_panorama_shell.send("set cli config-output-format set\n")  # Set the output format to (set) format
        ssh_panorama_shell.send("set cli scripting-mode on\n")
        ssh_panorama_shell.send("set cli pager off\n")
        ssh_panorama_shell.send("configure\n")  # Enter the configuration mode

        while True:
            if ssh_panorama_shell.recv_ready():
                _ = ssh_panorama_shell.recv(5 * 1024 * 1024).decode('utf-8')
                break
            elif ssh_panorama_shell.exit_status_ready():
                _ = ssh_panorama_shell.recv_exit_status()
                return False
            elif ssh_panorama_shell.closed or ssh_panorama_shell.eof_received or not ssh_panorama_shell.active:
                return False
            time.sleep(1)

        if vsys.lower() in ["", "any", "all"]:
            ssh_panorama_shell.send('show | match ip-netmask\\|ip-range\\|ip-wildcard\n')
        elif "," not in vsys:
            ssh_panorama_shell.send(f'show | match ip-netmask\\|ip-range\\|ip-wildcard | match "{vsys}"\n')
        elif "," in vsys:
            vsys = re.sub(",", r"\|", vsys)
            ssh_panorama_shell.send(f'show | match ip-netmask\\|ip-range\\|ip-wildcard | match "{vsys}"\n')
        ssh_panorama_shell.send(" " * 10)  # Send Spaces to show the all output

        # Receive the output
        # time.sleep(10)
        count = 0
        output = ""
        while True:
            if ssh_panorama_shell.recv_ready():  # Wait for the command to be ready
                count = 0
                output += ssh_panorama_shell.recv(10 * 1024 * 1024).decode('utf-8')
            elif ssh_panorama_shell.exit_status_ready():
                _ = ssh_panorama_shell.recv_exit_status()
                return False
            elif ssh_panorama_shell.closed or ssh_panorama_shell.eof_received or not ssh_panorama_shell.active:
                return False
            else:
                count += 1

            # Break the loop if no output is received for 5 seconds
            if count > 5:
                break

            time.sleep(1)

        # Clean the output
        output = re.sub(" +", " ", output)
        output = re.sub("\x1b.{,3}\x00*| *\x08|lines [0-9]+-[0-9]+[ \n]*", "", output)
        output = re.sub("\n *(?=[0-9]+)", " ", output)
        output = re.sub("\n+", "\n", output)
        output = re.sub("\n.et ", "\nset ", output)
        output = re.sub(r'(?m)^(?!set).*\n?', '', output)

        output = output.split('\n')
        for line in output:
            vsys, address_object, ip_address, ip_type = "", "", "", ""
            line = line.split(" ")
            line = [x.strip() for x in line]
            if len(line) > 3:
                if line[0] == "set" and line[1] == "device-group":
                    vsys = f"PAN - {line[2]}"
                    address_object = line[4]
                    ip_address = line[-1]
                elif line[0] == "set" and line[1] == "shared":
                    vsys = "PAN - shared"
                    address_object = line[3]
                    ip_address = line[-1]
                elif line[0] == "set" and line[1] == "address":
                    vsys = "PAN - shared"
                    address_object = line[2]
                    ip_address = line[-1]
                else:
                    continue
            else:
                continue

            if vsys and address_object and ip_address:
                # COMMENT THIS BLOCK IN PRODUCTION
                if address_object.lower().startswith("network_") and len(address_object) > 8:
                    address_object = address_object[8:]

                if get_subnet_type(address_object) != "Invalid":
                    continue
                # END

                ip_type = get_subnet_type(ip_address)
                if ip_type == "Invalid":
                    continue
                panorama_addresses.append([vsys, address_object, ip_address, ip_type])

        # DEBUG
        # for address in panorama_addresses:
        #     print(str(address))
        # END DEBUG

        if not len(panorama_addresses):
            set_info1("INFO", "green", "arrow")
            set_info2("No address objects are found in Palo Alto.", "red", "arrow")
            logs.set(f"INFO - No address objects are found in Palo Alto ({ip}).")
            return False

        panorama_checkbox_var.set(1)
        set_info1("INFO", "green", "arrow")
        set_info2(f"{str(len(panorama_addresses))} address objects has been exported from Palo Alto.", "green", "arrow")
        logs.set(f"INFO - {str(len(panorama_addresses))} address objects has been exported from Palo Alto ({ip}).")
        panorama_addresses_bkp = copy.deepcopy(panorama_addresses)
        for row in panorama_addresses:
            if isinstance(row[2], str):
                row[2] = convert_address_object(row[2], row[3])
        return True

    except Exception as e:
        set_info1("ERROR", "red", "arrow")
        set_info2("Failed exporting address objects from Palo Alto.", "red", "arrow")
        logs.set(f"ERROR - Failed exporting address objects from Palo Alto: {str(e)}")
        return False


def disconnect_from_panorama(logging=True):
    global ssh_panorama, logs, button_panorama, label_panorama_status

    # Disconnect from Panorama
    if ssh_panorama is not None or isinstance(ssh_panorama, paramiko.SSHClient):
        try:
            ssh_panorama.close()
        finally:
            ssh_panorama = None

    ssh_panorama = None

    # Update the Panorama status label
    button_panorama.config(text="Connect")
    label_panorama_status.config(text="Not Connected", foreground="red")
    if logging:
        logs.set("INFO - Disconnected from Palo Alto Device.")
    enable_buttons()


def get_panorama_credentials():
    global gui, panorama_ip, panorama_username, panorama_password, panorama_vsys

    def close_panorama_credentials_window():
        global gui
        gui.attributes('-disabled', False)
        panorama_credentials_window.destroy()
        gui.update()
        gui.focus_force()

    # Create a save button
    def save_panorama_credentials(remember):
        global panorama_ip, panorama_username, panorama_password, panorama_vsys, KEY

        # Prevent the user from clicking the save button multiple times
        save_button.config(state="disabled", cursor="arrow")
        cancel_button.config(state="disabled", cursor="arrow")
        panorama_credentials_window.protocol("WM_DELETE_WINDOW", lambda: False)

        # Set the Panorama credentials
        panorama_ip_tmp = panorama_ip_entry.get().strip()
        panorama_username_tmp = panorama_username_entry.get().strip()
        panorama_password_tmp = panorama_password_entry.get()
        panorama_vsys_tmp = panorama_vsys_entry.get().strip()

        # Check if the Panorama credentials are correct
        if not connect_to_panorama(panorama_ip_tmp,
                                   panorama_username_tmp,
                                   panorama_password_tmp,
                                   panorama_vsys_tmp,
                                   panorama_credentials_window):

            # Check if the user wants to continue
            if not tk.messagebox.askyesno("Warning",
                                          "Can not connect to the Palo Alto Device.\nSaving these credentials?",
                                          icon="warning",
                                          default="no",
                                          parent=panorama_credentials_window):
                save_button.config(state="normal", cursor="hand2")
                cancel_button.config(state="normal", cursor="hand2")
                panorama_credentials_window.protocol("WM_DELETE_WINDOW", lambda: close_panorama_credentials_window())
                return False

        # Set the Panorama credentials
        global panorama_ip, panorama_username, panorama_password, panorama_vsys
        panorama_ip = panorama_ip_tmp
        panorama_username = panorama_username_tmp
        panorama_password = panorama_password_tmp
        panorama_vsys = panorama_vsys_tmp

        # Check if the user wants to save the Panorama credentials
        if remember == 1:
            global KEY
            # Save the Panorama credentials to a file
            credentials = f"PAN\n{panorama_ip}\n{panorama_username}\n{panorama_password}\n{panorama_vsys}"

            cipher_suite = Fernet(KEY)
            cipher_text = cipher_suite.encrypt(credentials.encode('utf-8'))
            with open('panorama.cfg', 'w') as f:
                f.write(cipher_text.decode('utf-8'))
            logs.set("INFO - Palo Alto Device credentials are saved to a file (panorama.cfg).")

        # Close the Panorama credentials window
        close_panorama_credentials_window()

    if panorama_ip == "" and panorama_username == "" and panorama_password == "":
        import_credentials("panorama")

    # Create a Panorama credentials window
    panorama_credentials_window = tk.Toplevel(gui)
    panorama_credentials_window.title('PAN Credentials')
    panorama_credentials_window.focus_force()
    panorama_credentials_window.grab_set()
    panorama_credentials_window.focus_set()
    panorama_credentials_window.transient(gui)
    gui.attributes('-disabled', True)

    panorama_credentials_window.geometry('%dx%d+%d+%d' % (350, 230, gui.winfo_rootx() + 100, gui.winfo_rooty() + 75))
    panorama_credentials_window.resizable(False, False)
    panorama_credentials_window.iconphoto(False, icon)

    # Create a frame for the Panorama credentials window
    panorama_credentials_frame = tk.Frame(panorama_credentials_window)
    panorama_credentials_frame.pack()

    # Create a label for the Panorama credentials window
    panorama_credentials_label = ttk.Label(panorama_credentials_frame, text='Palo Alto Credentials',
                                           font=("Arial bold", 10))
    panorama_credentials_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    # Create a label for the Panorama IP address
    panorama_ip_label = ttk.Label(panorama_credentials_frame, text='Panorama IP:', font=("Arial bold", 8))
    panorama_ip_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')

    # Create an entry for the Panorama IP address
    panorama_ip_entry = ttk.Entry(panorama_credentials_frame, width=37)
    panorama_ip_entry.bind("<Button-3>", right_click)
    panorama_ip_entry.insert(0, panorama_ip)
    panorama_ip_entry.grid(row=1, column=1, padx=5, pady=5)
    panorama_ip_entry.focus_set()

    # Create a label for the Panorama username
    panorama_username_label = ttk.Label(panorama_credentials_frame, text='Username:', font=("Arial bold", 8))
    panorama_username_label.grid(row=2, column=0, padx=5, pady=5, sticky='e')

    # Create an entry for the Panorama username
    panorama_username_entry = ttk.Entry(panorama_credentials_frame, width=37)
    panorama_username_entry.bind("<Button-3>", right_click)
    panorama_username_entry.insert(0, panorama_username)
    panorama_username_entry.grid(row=2, column=1, padx=5, pady=5)

    # Create a label for the Panorama password
    panorama_password_label = ttk.Label(panorama_credentials_frame, text='Password:', font=("Arial bold", 8))
    panorama_password_label.grid(row=3, column=0, padx=5, pady=5, sticky='e')

    # Create an entry for the Panorama password
    panorama_password_entry = ttk.Entry(panorama_credentials_frame, show='*', width=37)
    panorama_password_entry.bind("<Button-3>", right_click)
    panorama_password_entry.insert(0, panorama_password)
    panorama_password_entry.grid(row=3, column=1, padx=5, pady=5)

    # Create a label for the Panorama vsys
    panorama_vsys_label = ttk.Label(panorama_credentials_frame, text='Vsys:', font=("Arial", 8))
    panorama_vsys_label.grid(row=4, column=0, padx=5, pady=5, sticky='e')

    # Create an entry for the Panorama vsys
    panorama_vsys_entry = ttk.Entry(panorama_credentials_frame, width=37)
    panorama_vsys_entry.bind("<Button-3>", right_click)
    panorama_vsys_entry.insert(0, panorama_vsys)
    panorama_vsys_entry.grid(row=4, column=1, padx=5, pady=5)

    # Create a remember me checkbutton
    remember_me = tk.IntVar()
    remember_me_checkbutton = ttk.Checkbutton(panorama_credentials_frame, text='Remember Me?', variable=remember_me,
                                              onvalue=1, offvalue=0, cursor='hand2', takefocus=False, width=15)
    remember_me_checkbutton.grid(row=5, column=0, columnspan=2, padx=25, pady=5, sticky='w')

    # Create a frame for the buttons
    buttons_frame = ttk.Frame(panorama_credentials_frame)
    buttons_frame.grid(row=6, column=0, columnspan=2, padx=5, pady=5)

    # Create a save button
    save_button = ttk.Button(buttons_frame, text='Save', cursor='hand2', width=15,
                             command=lambda: save_panorama_credentials(remember_me.get()))
    save_button.grid(row=6, column=0, padx=5, pady=5, sticky='e')

    # Create a cancel button
    cancel_button = ttk.Button(buttons_frame, text='Cancel', cursor='hand2', width=15,
                               command=close_panorama_credentials_window)
    cancel_button.grid(row=6, column=1, padx=5, pady=5, sticky='w')

    panorama_credentials_window.protocol("WM_DELETE_WINDOW", close_panorama_credentials_window)
    panorama_credentials_window.bind('<Return>', lambda event: save_panorama_credentials(remember_me.get()))

    panorama_credentials_window.mainloop()


################################################################

def connect_to_forti(ip=forti_ip, port=forti_port, username=forti_username, password=forti_password, vdom=forti_vdom,
                     window=None):
    global forti_api, logs, start_flag, forti_addresses, forti_ip, forti_username, forti_password, forti_vdom

    disable_buttons()

    if forti_api is not None:
        disconnect_from_forti(False)

    if ip.strip() == "" or port.strip() == "" or username.strip() == "" or password == "":
        button_forti.config(text="Connect")
        enable_buttons()
        disconnect_from_forti(False)
        logs.set("ERROR - Fortinet credentials are empty.")
        return False

    button_forti.config(cursor="wait", text="Connecting")
    try:
        forti_api = FortiGateAPI(host=ip,
                                 port=port,
                                 username=username,
                                 password=password,
                                 vdom=vdom,
                                 timeout=5,
                                 scheme="https")
        forti_api.login()
    except:
        if start_flag:
            logs.set(
                f"ERROR - Failed connecting to Fortinet device using https protocol on port {port}, trying to connect using http protocol.")
        try:
            forti_api = FortiGateAPI(host=ip,
                                     port=port,
                                     username=username,
                                     password=password,
                                     vdom=vdom,
                                     timeout=5,
                                     scheme="http")
            forti_api.login()
        except:

            button_forti.config(text="Connect")
            enable_buttons()
            if start_flag:
                logs.set(f"ERROR - Failed connecting to Fortinet device using http protocol on port {port}.")
                disconnect_from_forti()
            else:
                logs.set(f"ERROR - Failed connecting to Fortinet: timed out")
                disconnect_from_forti(False)
            return False

    label_forti_status.config(text="Connected", foreground="green")
    logs.set(f"INFO - Connected to Fortinet Device ({ip}).")

    try:
        export_forti_thread = PropagatingThread(target=export_forti_address_objects,
                                                args=(ip, vdom),
                                                daemon=True,
                                                name="export_forti_thread")
        export_forti_thread.start()

        while export_forti_thread.is_alive():
            gui.update()
            if window is not None:
                window.update()
            time.sleep(0.1)

        export_forti_thread.join()
    except Exception as e:
        forti_addresses = []
        logs.set(f"ERROR - Failed exporting address objects from Fortinet: {str(e)}")

    button_forti.config(text="Connect")
    enable_buttons()
    return True


def export_forti_address_objects(ip=forti_ip, vdom=forti_vdom, window=None):
    global forti_api, separators, forti_addresses, logs, forti_addresses_bkp

    if forti_api is None:
        return False

    object_keys = ['subnet', 'start-ip', 'end-ip']

    try:
        forti_addresses = []
        vdoms = [v["name"] for v in forti_api.cmdb.system.vdom.get()]

        if vdom.lower() not in ["", "all"]:
            if any(sep in vdom for sep in separators):
                sep = [sep for sep in separators if sep in vdom][0]
                input_vdoms = [v.strip() for v in vdom.split(sep)]
            else:
                input_vdoms = [vdom]

            tmp_vdoms = []
            for v in input_vdoms:
                if v in vdoms:
                    tmp_vdoms.append(v)
                else:
                    logs.set(f"ERROR - VDOM ({v}) is not found in the Fortinet Device {forti_ip}.")
            vdoms = tmp_vdoms.copy()

        items = []
        for vdom in vdoms:
            forti_api.vdom = vdom
            items_tmp = forti_api.cmdb.firewall.address.get()
            for i in range(len(items_tmp)):
                items_tmp[i]["vdom"] = vdom
            items.extend(items_tmp)
            del items_tmp

        for i in range(len(items)):
            if any(items[i].get(key) for key in object_keys):
                if "subnet" in items[i].keys():
                    ip_netmask = items[i]["subnet"].strip().split(" ")
                    if len(ip_netmask) == 2:
                        items[i]["subnet"] = "/".join([ip_netmask[0].strip(),
                                                       str(ipaddress.ip_network(
                                                           f"0.0.0.0/{ip_netmask[1].strip()}").prefixlen)])
                if "start-ip" in items[i].keys() and "end-ip" in items[i].keys():
                    items[i]["subnet"] = f"{items[i]['start-ip'].strip()}-{items[i]['end-ip'].strip()}"

        for i in range(len(items) - 1, -1, -1):
            if "subnet" not in items[i].keys():
                items.pop(i)
            elif items[i]["subnet"] in ["0.0.0.0/32", "0.0.0.0/0"]:
                items.pop(i)

            if items[i]["name"].lower().startswith("network_") and len(items[i]["name"]) > 8:
                items[i]["name"] = items[i]["name"][8:]

            if get_subnet_type(items[i]["name"]) != "Invalid":
                items.pop(i)

        forti_addresses = []
        for item in items:
            forti_addresses.append(
                [f"FGT - {item["vdom"]}", item["name"], item["subnet"], get_subnet_type(item["subnet"])])

        if not len(forti_addresses):
            set_info1("INFO", "green", "arrow")
            set_info2("No address objects are found in Fortinet.", "red", "arrow")
            logs.set(f"INFO - No address objects are found in Fortinet ({ip}).")
            return False

        forti_checkbox_var.set(1)
        set_info1("INFO", "green", "arrow")
        set_info2(f"{str(len(forti_addresses))} address objects has been exported from Fortinet.", "green", "arrow")
        logs.set(f"INFO - {str(len(forti_addresses))} address objects has been exported from Fortinet ({ip}).")
        forti_addresses_bkp = copy.deepcopy(forti_addresses)
        for row in forti_addresses:
            if isinstance(row[2], str):
                row[2] = convert_address_object(row[2], row[3])
        return True

    except Exception as e:
        set_info1("ERROR", "red", "arrow")
        set_info2("Failed exporting address objects from Fortinet.", "red", "arrow")
        logs.set(f"ERROR - Failed exporting address objects from Fortinet: {str(e)}")
        return False


def disconnect_from_forti(logging=True):
    global forti_api, logs

    if forti_api is not None:
        try:
            forti_api.logout()
        finally:
            forti_api = None

    forti_api = None

    button_forti.config(text="Connect")
    label_forti_status.config(text="Not Connected", foreground="red")
    if logging:
        logs.set("INFO - Disconnected from Fortinet Device.")
    enable_buttons()


def get_forti_credentials():
    global forti_ip, forti_port, forti_username, forti_password, forti_vdom

    def close_fortigate_credentials_window():
        gui.attributes('-disabled', False)
        fortigate_credentials_window.destroy()
        gui.update()
        gui.focus_force()

    def save_fortigate_credentials(remember):
        global logs

        save_button.config(state="disabled", cursor="arrow")
        cancel_button.config(state="disabled", cursor="arrow")
        fortigate_credentials_window.protocol("WM_DELETE_WINDOW", lambda: False)

        fortigate_ip_tmp = fortigate_ip_entry.get().strip()
        forti_port_tmp = fortigate_port_entry.get().strip()
        fortigate_username_tmp = fortigate_username_entry.get().strip()
        fortigate_password_tmp = fortigate_password_entry.get()
        fortigate_vdom_tmp = fortigate_vdom_entry.get().strip()

        connection = PropagatingThread(target=connect_to_forti,
                                       args=(
                                           fortigate_ip_tmp, forti_port_tmp, fortigate_username_tmp,
                                           fortigate_password_tmp,
                                           fortigate_vdom_tmp, fortigate_credentials_window),
                                       daemon=True,
                                       name="fortigate_connection_thread")
        connection.start()
        while connection.is_alive():
            gui.update()
            time.sleep(0.1)
        connection = connection.join()

        if not connection:
            if not tk.messagebox.askyesno("Warning",
                                          "Can not connect to the Fortinet Device.\nSaving these credentials?",
                                          icon="warning", default="no", parent=fortigate_credentials_window):
                save_button.config(state="normal", cursor="hand2")
                cancel_button.config(state="normal", cursor="hand2")
                fortigate_credentials_window.protocol("WM_DELETE_WINDOW", lambda: close_fortigate_credentials_window())
                return False

        global forti_ip, forti_port, forti_username, forti_password, forti_vdom
        forti_ip = fortigate_ip_tmp
        forti_port = forti_port_tmp
        forti_username = fortigate_username_tmp
        forti_password = fortigate_password_tmp
        forti_vdom = fortigate_vdom_tmp

        if remember == 1:
            global KEY
            credentials = f"FORTI\n{forti_ip}\n{forti_port}\n{forti_username}\n{forti_password}\n{forti_vdom}"

            cipher_suite = Fernet(KEY)
            cipher_text = cipher_suite.encrypt(credentials.encode('utf-8'))
            with open('fortigate.cfg', 'w') as f:
                f.write(cipher_text.decode('utf-8'))
            logs.set("INFO - Fortinet Device credentials are saved to a file (fortigate.cfg).")

        close_fortigate_credentials_window()

    if forti_ip == "" and forti_username == "" and forti_password == "":
        import_credentials("forti")

    fortigate_credentials_window = tk.Toplevel(gui)
    fortigate_credentials_window.title('Fortinet Credentials')
    fortigate_credentials_window.focus_force()
    fortigate_credentials_window.grab_set()
    fortigate_credentials_window.focus_set()
    fortigate_credentials_window.transient(gui)
    gui.attributes('-disabled', True)

    fortigate_credentials_window.geometry('%dx%d+%d+%d' % (350, 260, gui.winfo_rootx() + 100, gui.winfo_rooty() + 75))
    fortigate_credentials_window.resizable(False, False)
    fortigate_credentials_window.iconphoto(False, icon)

    fortigate_credentials_frame = tk.Frame(fortigate_credentials_window)
    fortigate_credentials_frame.pack()

    fortigate_credentials_label = ttk.Label(fortigate_credentials_frame, text='Fortinet Credentials',
                                            font=("Arial bold", 10))
    fortigate_credentials_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    fortigate_ip_label = ttk.Label(fortigate_credentials_frame, text='Forti IP:', font=("Arial bold", 8))
    fortigate_ip_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')

    fortigate_ip_entry = ttk.Entry(fortigate_credentials_frame, width=37)
    fortigate_ip_entry.bind("<Button-3>", right_click)
    fortigate_ip_entry.insert(0, forti_ip)
    fortigate_ip_entry.grid(row=1, column=1, padx=5, pady=5)
    fortigate_ip_entry.focus_set()

    fortigate_port_label = ttk.Label(fortigate_credentials_frame, text='Port:', font=("Arial bold", 8))
    fortigate_port_label.grid(row=2, column=0, padx=5, pady=5, sticky='e')

    fortigate_port_entry = ttk.Entry(fortigate_credentials_frame, width=37)
    fortigate_port_entry.bind("<Button-3>", right_click)
    fortigate_port_entry.insert(0, forti_port)
    fortigate_port_entry.grid(row=2, column=1, padx=5, pady=5)

    fortigate_username_label = ttk.Label(fortigate_credentials_frame, text='Username:', font=("Arial bold", 8))
    fortigate_username_label.grid(row=3, column=0, padx=5, pady=5, sticky='e')

    fortigate_username_entry = ttk.Entry(fortigate_credentials_frame, width=37)
    fortigate_username_entry.bind("<Button-3>", right_click)
    fortigate_username_entry.insert(0, forti_username)
    fortigate_username_entry.grid(row=3, column=1, padx=5, pady=5)

    fortigate_password_label = ttk.Label(fortigate_credentials_frame, text='Password:', font=("Arial bold", 8))
    fortigate_password_label.grid(row=4, column=0, padx=5, pady=5, sticky='e')

    fortigate_password_entry = ttk.Entry(fortigate_credentials_frame, show='*', width=37)
    fortigate_password_entry.bind("<Button-3>", right_click)
    fortigate_password_entry.insert(0, forti_password)
    fortigate_password_entry.grid(row=4, column=1, padx=5, pady=5)

    fortigate_vdom_label = ttk.Label(fortigate_credentials_frame, text='VDOM:', font=("Arial", 8))
    fortigate_vdom_label.grid(row=5, column=0, padx=5, pady=5, sticky='e')

    fortigate_vdom_entry = ttk.Entry(fortigate_credentials_frame, width=37)
    fortigate_vdom_entry.bind("<Button-3>", right_click)
    fortigate_vdom_entry.insert(0, forti_vdom)
    fortigate_vdom_entry.grid(row=5, column=1, padx=5, pady=5)

    remember_me = tk.IntVar()
    remember_me_checkbutton = ttk.Checkbutton(fortigate_credentials_frame, text='Remember Me?', variable=remember_me,
                                              onvalue=1, offvalue=0, cursor='hand2', takefocus=False, width=15)
    remember_me_checkbutton.grid(row=6, column=0, columnspan=2, padx=25, pady=5, sticky='w')

    buttons_frame = ttk.Frame(fortigate_credentials_frame)
    buttons_frame.grid(row=7, column=0, columnspan=2, padx=5, pady=5)

    save_button = ttk.Button(buttons_frame, text='Save', cursor='hand2', width=15,
                             command=lambda: save_fortigate_credentials(remember_me.get()))
    save_button.grid(row=0, column=0, padx=5, pady=5, sticky='e')

    cancel_button = ttk.Button(buttons_frame, text='Cancel', cursor='hand2', width=15,
                               command=lambda: close_fortigate_credentials_window())
    cancel_button.grid(row=0, column=1, padx=5, pady=5, sticky='w')

    fortigate_credentials_window.protocol("WM_DELETE_WINDOW", lambda: close_fortigate_credentials_window())
    fortigate_credentials_window.bind('<Return>', lambda event: save_fortigate_credentials(remember_me.get()))

    fortigate_credentials_window.mainloop()


################################################################

def connect_to_aci(ip=aci_ip, username=aci_username, password=aci_password, apic_class=aci_class, window=None):
    global ssh_aci, logs, aci_ip, aci_username, aci_password, aci_class, aci_addresses

    disable_buttons()

    if ssh_aci is not None:
        disconnect_from_aci()

    # Check if the ACI credentials are correct
    if ip == "" or username == "" or password == "":
        button_aci.config(text="Connect")
        enable_buttons()
        disconnect_from_aci(False)
        logs.set("ERROR - ACI credentials are empty.")
        return False

    # Check if the ACI credentials are correct
    ssh_aci = paramiko.SSHClient()
    ssh_aci.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    button_aci.config(cursor="wait", text="Connecting")
    try:
        # Connect to ACI
        ssh_aci_thread = PropagatingThread(target=ssh_aci.connect,
                                           args=(ip, 22, username, password),
                                           kwargs={'timeout': 10,
                                                   'allow_agent': False,
                                                   'look_for_keys': False},
                                           daemon=True,
                                           name="ssh_aci_thread")
        ssh_aci_thread.start()

        while ssh_aci_thread.is_alive():
            gui.update()
            if window is not None:
                window.update()
            time.sleep(0.1)

        ssh_aci_thread.join()

        # Check if the connection is still open
        if ssh_aci.get_transport().is_active():
            try:
                transport = ssh_aci.get_transport()
                transport.send_ignore()
            except Exception as e:
                logs.set(f"ERROR - Failed connecting to APIC: {str(e)}")
                button_aci.config(text="Connect")
                enable_buttons()
                disconnect_from_aci()
                return False

        # Update the ACI status label
        label_aci_status.config(text="Connected", foreground="green")
        logs.set(f"INFO - Connected to APIC ({ip}).")

        try:
            if isinstance(ssh_aci, paramiko.SSHClient) and ssh_aci.get_transport().is_active():
                export_aci_thread = PropagatingThread(target=export_aci_address_objects,
                                                      args=(ip, apic_class),
                                                      daemon=True,
                                                      name="export_aci_thread")
                export_aci_thread.start()

                while export_aci_thread.is_alive():
                    gui.update()
                    if window is not None:
                        window.update()
                    time.sleep(0.1)

                export_aci_thread.join()
        except Exception as e:
            aci_addresses = []
            logs.set(f"ERROR - Failed exporting address objects from APIC: {str(e)}")

        button_aci.config(text="Connect")
        enable_buttons()
        return True
    except Exception as e:
        logs.set(f"ERROR - Failed connecting to APIC: {str(e)}")
        button_aci.config(text="Connect")
        enable_buttons()
        disconnect_from_aci(False)
        return False


def export_aci_address_objects(ip=aci_ip, apic_class=aci_class):
    global ssh_aci, aci_ip, aci_username, aci_password, aci_class, aci_addresses, logs, aci_addresses_bkp

    try:
        aci_addresses = []
        ip = ip.strip()
        apic_class = apic_class.strip()

        if ssh_aci is None:
            return False

        elif isinstance(ssh_aci, paramiko.SSHClient):
            stdin, stdout, stderr = ssh_aci.exec_command(f'moquery -c {apic_class}' +
                                                         r' | grep -E "^dn[ \t]*:.+\[.*[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+.*\]" | cut -d ":" -f 2 | cut -d " " -f 2')
            output = stdout.read().decode('utf-8')

            output = output.split('\n')
            for line in output:
                tenant, address_object, ip_address, ip_type = "", "", "", ""
                line = line.strip().split("/")
                line = [x.strip() for x in line]
                if re.match("[0-3][0-9]]", line[-1]):
                    line[-2] = f"{line[-2]}/{line[-1]}"
                    line = line[:-1]

                if len(line) == 4:
                    if line[1].lower().startswith("tn-"):
                        tenant = f"ACI - {line[1][3:]}"
                    else:
                        tenant = f"ACI - {line[1]}"

                    if line[2].upper().startswith("BD-"):
                        address_object = line[2][3:]
                    else:
                        address_object = line[2]

                    if line[-1].lower().startswith("subnet-["):
                        ip_address = line[3][8:-1]

                elif len(line) == 5:
                    if line[1].lower().startswith("tn-"):
                        tenant = "ACI - " + line[1][3:]
                    else:
                        tenant = "ACI - " + line[1]

                    if line[2].lower().startswith("ap-") and line[3].lower().startswith("epg-"):
                        address_object = f"{line[2][3:]}-{line[3][4:]}"
                    else:
                        address_object = f"{line[2]}-{line[3]}"

                    if line[-1].lower().startswith("subnet-["):
                        ip_address = line[-1][8:-1]

                if tenant and address_object and ip_address:
                    # COMMENT THIS BLOCK IN PRODUCTION
                    if get_subnet_type(address_object) != "Invalid":
                        continue
                    # END

                    ip_type = get_subnet_type(ip_address)
                    if ip_type == "Invalid":
                        continue

                    aci_addresses.append([tenant, f"{tenant[6:]}-{address_object}", ip_address, ip_type])

            # DEBUG
            # for address in aci_addresses:
            #     print(str(address))
            # END DEBUG

            if not len(aci_addresses):
                set_info1("INFO", "green", "arrow")
                set_info2("No address objects are found in APIC.", "red", "arrow")
                logs.set(f"INFO - No address objects are found in APIC ({ip}).")
                return False

            aci_checkbox_var.set(1)
            set_info1("INFO", "green", "arrow")
            set_info2(f"{str(len(aci_addresses))} address objects has been exported from APIC.", "green", "arrow")
            logs.set(f"INFO - {str(len(aci_addresses))} address objects has been exported from APIC ({ip}).")
            aci_addresses_bkp = copy.deepcopy(aci_addresses)
            for row in aci_addresses:
                if isinstance(row[2], str):
                    row[2] = convert_address_object(row[2], row[3])
            return True

    except Exception as e:
        set_info1("ERROR", "red", "arrow")
        set_info2("Failed exporting address objects from APIC server.", "red", "arrow")
        logs.set(f"ERROR - Failed exporting address objects from APIC server: {str(e)}")
        return False


def disconnect_from_aci(logging=True):
    global ssh_aci, logs

    # Disconnect from ACI
    if ssh_aci is not None or isinstance(ssh_aci, paramiko.SSHClient):
        try:
            ssh_aci.close()
        finally:
            ssh_aci = None

    ssh_aci = None

    # Update the ACI status label
    button_aci.config(text="Connect")
    label_aci_status.config(text="Not Connected", foreground="red")
    if logging:
        logs.set("INFO - Disconnected from ACI.")
    enable_buttons()


def get_aci_credentials():
    global aci_ip, aci_username, aci_password, aci_class

    def close_aci_credentials_window():
        gui.attributes('-disabled', False)
        aci_credentials_window.destroy()
        gui.update()
        gui.focus_force()

    # Create a save button
    def save_aci_credentials(remember):
        global logs

        # Prevent the user from clicking the save button multiple times
        save_button.config(state="disabled", cursor="arrow")
        cancel_button.config(state="disabled", cursor="arrow")
        aci_credentials_window.protocol("WM_DELETE_WINDOW", lambda: False)

        # Set the ACI credentials
        aci_ip_tmp = aci_ip_entry.get().strip()
        aci_username_tmp = aci_username_entry.get().strip()
        aci_password_tmp = aci_password_entry.get()
        aci_class_tmp = aci_class_entry.get().strip()

        if aci_class_tmp == "" or " " in aci_class_tmp:
            tk.messagebox.showerror("Error",
                                    "Class(es) can not be empty or include spaces.\nOnly comma separated classes are allowed.",
                                    parent=aci_credentials_window)

            save_button.config(state="normal", cursor="hand2")
            cancel_button.config(state="normal", cursor="hand2")
            aci_credentials_window.protocol("WM_DELETE_WINDOW", lambda: close_aci_credentials_window())
            return False

        # Check if the ACI credentials are correct
        if not connect_to_aci(aci_ip_tmp,
                              aci_username_tmp,
                              aci_password_tmp,
                              aci_class_tmp,
                              aci_credentials_window):

            # Check if the user wants to continue
            if not tk.messagebox.askyesno("Warning",
                                          "Can not connect to the APIC server.\nSaving these credentials?",
                                          icon="warning",
                                          default="no",
                                          parent=aci_credentials_window):
                save_button.config(state="normal", cursor="hand2")
                cancel_button.config(state="normal", cursor="hand2")
                aci_credentials_window.protocol("WM_DELETE_WINDOW", lambda: close_aci_credentials_window())
                return

        # Set the ACI credentials
        global aci_ip, aci_username, aci_password, aci_class, aci_addresses
        aci_ip = aci_ip_tmp
        aci_username = aci_username_tmp
        aci_password = aci_password_tmp
        aci_class = aci_class_tmp

        # Check if the user wants to save the ACI credentials
        if remember == 1:
            global KEY

            credentials = f"ACI\n{aci_ip}\n{aci_username}\n{aci_password}\n{aci_class}"

            cipher_suite = Fernet(KEY)
            cipher_text = cipher_suite.encrypt(credentials.encode('utf-8'))

            # Save the ACI credentials to a file
            with open('aci.cfg', 'w') as f:
                f.write(cipher_text.decode('utf-8'))
            logs.set("INFO - ACI credentials are saved to a file (aci.cfg).")

        # Close the ACI credentials window
        close_aci_credentials_window()

    if aci_ip == "" and aci_username == "" and aci_password == "":
        import_credentials("aci")

    # Create an ACI credentials window
    aci_credentials_window = tk.Toplevel(gui)
    aci_credentials_window.title('CISCO ACI Credentials')
    aci_credentials_window.focus_force()
    aci_credentials_window.grab_set()
    aci_credentials_window.focus_set()
    aci_credentials_window.transient(gui)
    gui.attributes('-disabled', True)

    aci_credentials_window.geometry('%dx%d+%d+%d' % (350, 235, gui.winfo_rootx() + 100, gui.winfo_rooty() + 70))
    aci_credentials_window.resizable(False, False)
    aci_credentials_window.iconphoto(False, icon)

    # Create a frame for the ACI credentials window
    aci_credentials_frame = tk.Frame(aci_credentials_window)
    aci_credentials_frame.pack()

    # Create a label for the ACI credentials window
    aci_credentials_label = ttk.Label(aci_credentials_frame, text='APIC Credentials', font=("Arial bold", 10))
    aci_credentials_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    # Create a label for the APIC IP address
    aci_ip_label = ttk.Label(aci_credentials_frame, text='APIC IP:', font=("Arial bold", 8))
    aci_ip_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')

    # Create an entry for the APIC IP address
    aci_ip_entry = ttk.Entry(aci_credentials_frame, width=40)
    aci_ip_entry.bind("<Button-3>", right_click)
    aci_ip_entry.insert(0, aci_ip)
    aci_ip_entry.grid(row=1, column=1, padx=5, pady=5)

    # Create a label for the ACI username
    aci_username_label = ttk.Label(aci_credentials_frame, text='Username:', font=("Arial bold", 8))
    aci_username_label.grid(row=2, column=0, padx=5, pady=5, sticky='e')

    # Create an entry for the ACI username
    aci_username_entry = ttk.Entry(aci_credentials_frame, width=40)
    aci_username_entry.bind("<Button-3>", right_click)
    aci_username_entry.insert(0, aci_username)
    aci_username_entry.grid(row=2, column=1, padx=5, pady=5)

    # Create a label for the ACI password
    aci_password_label = ttk.Label(aci_credentials_frame, text='Password:', font=("Arial bold", 8))
    aci_password_label.grid(row=3, column=0, padx=5, pady=5, sticky='e')

    # Create an entry for the ACI password
    aci_password_entry = ttk.Entry(aci_credentials_frame, show='*', width=40)
    aci_password_entry.bind("<Button-3>", right_click)
    aci_password_entry.insert(0, aci_password)
    aci_password_entry.grid(row=3, column=1, padx=5, pady=5)

    # Create a label for the ACI Class / DN
    aci_class_label = ttk.Label(aci_credentials_frame, text='Class:', font=("Arial bold", 8))
    aci_class_label.grid(row=4, column=0, padx=5, pady=5, sticky='e')

    # Create an entry for the ACI Class / DN
    aci_class_entry = ttk.Entry(aci_credentials_frame, width=40)
    aci_class_entry.bind("<Button-3>", right_click)
    aci_class_entry.insert(0, aci_class)
    aci_class_entry.grid(row=4, column=1, padx=5, pady=5)

    # Create a remember me checkbutton
    remember_me = tk.IntVar()
    remember_me_checkbutton = ttk.Checkbutton(aci_credentials_frame, text='Remember Me?', variable=remember_me,
                                              onvalue=1, offvalue=0, cursor='hand2', takefocus=False, width=15)
    remember_me_checkbutton.grid(row=5, column=0, columnspan=2, padx=25, pady=5, sticky='w')

    # Create a frame for the buttons
    buttons_frame = ttk.Frame(aci_credentials_frame)
    buttons_frame.grid(row=7, column=0, columnspan=2, padx=5, pady=5)

    # Create a save button
    save_button = ttk.Button(buttons_frame, text='Save', cursor='hand2', width=15,
                             command=lambda: save_aci_credentials(remember_me.get()))
    save_button.grid(row=7, column=0, padx=5, pady=5, sticky='e')

    # Create a cancel button
    cancel_button = ttk.Button(buttons_frame, text='Cancel', cursor='hand2', width=15,
                               command=lambda: close_aci_credentials_window())
    cancel_button.grid(row=7, column=1, padx=5, pady=5, sticky='w')

    aci_credentials_window.protocol("WM_DELETE_WINDOW", lambda: close_aci_credentials_window())
    aci_credentials_window.bind('<Return>', lambda event: save_aci_credentials(remember_me.get()))

    aci_credentials_window.mainloop()


################################################################

def check_dns_servers(*servers):
    global flags, dns_checkbox_var, logs, resolvers, dns_resolver, dns_servers, button_dns, label_dns_status
    if isinstance(servers, tuple) and len(servers) == 1 and isinstance(servers[0], list):
        servers = servers[0]
    else:
        servers = list(servers)

    disable_buttons()
    button_dns.config(cursor="wait", text="Checking")
    output = []
    for server in servers:
        if get_subnet_type(server) != "IP":
            continue

        try:
            if ipaddress.ip_address(server).version != 4:
                continue
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(2.25)
            result = sock.connect_ex((server, 53))
            if result == 0:
                output.append(server)
            sock.close()
        except:
            pass

    if not len(output):
        label_dns_status.config(text="No Reachable DNS Servers", foreground="red")
        logs.set("ERROR - No reachable DNS servers are found.")
        dns_checkbox_var.set(0)
        flags[4] = False
    else:
        label_dns_status.config(text=f"{len(output)} Reachable DNS Servers", foreground="green")
        logs.set(f"INFO - {len(output)} reachable DNS servers are found.")
        dns_checkbox_var.set(1)

        dns_resolver = dns.resolver.Resolver()
        resolvers = output.copy()
        dns_resolver.nameservers = resolvers

        flags[4] = True

    button_dns.config(cursor="arrow", text="Check")
    enable_buttons()
    return output


def get_dns_servers():
    global resolvers, dns_servers, dns_checkbox_var, logs

    def close_dns_window():
        global gui
        gui.attributes('-disabled', False)
        dns_window.destroy()
        gui.update()
        gui.focus_force()

    def save_dns_servers(remember):
        global logs, dns_servers, resolvers, KEY

        save_button.config(state="disabled", cursor="arrow")
        cancel_button.config(state="disabled", cursor="arrow")
        dns_window.protocol("WM_DELETE_WINDOW", lambda: False)

        dns_servers_tmp = [dns_svr1_entry.get().strip(),
                           dns_svr2_entry.get().strip(),
                           dns_svr3_entry.get().strip(),
                           dns_svr4_entry.get().strip()]

        for i in range(3, -1, -1):
            if dns_servers_tmp[i] == "":
                dns_servers_tmp.pop(i)
        if len(dns_servers_tmp) < 4:
            dns_servers_tmp.extend([""] * (4 - len(dns_servers_tmp)))

        check = PropagatingThread(target=check_dns_servers,
                                  args=(dns_servers_tmp,),
                                  daemon=True,
                                  name="check_dns_servers_thread")
        check.start()
        while check.is_alive():
            gui.update()
            time.sleep(0.1)

        servers = check.join()

        invalids_server = []
        for server in dns_servers_tmp:
            if server not in servers and server != "":
                invalids_server.append(server)

        if not len(servers):
            if not tk.messagebox.askyesno("Error",
                                          "No reachable DNS servers are found.\nSaving these servers?",
                                          icon="error",
                                          default="no",
                                          parent=dns_window):
                save_button.config(state="normal", cursor="hand2")
                cancel_button.config(state="normal", cursor="hand2")
                dns_window.protocol("WM_DELETE_WINDOW", lambda: close_dns_window())
                return False
        elif len(invalids_server):
            if not tk.messagebox.askyesno("Warning",
                                          f"Invalid DNS servers are found.\n{"\n".join(invalids_server)}\n\nSaving these servers?",
                                          icon="warning",
                                          default="no",
                                          parent=dns_window):
                save_button.config(state="normal", cursor="hand2")
                cancel_button.config(state="normal", cursor="hand2")
                dns_window.protocol("WM_DELETE_WINDOW", lambda: close_dns_window())
                return False

        dns_servers = dns_servers_tmp.copy()
        resolvers = servers.copy()

        if remember == 1:
            global KEY
            credentials = f"DNS\n{"\n".join(dns_servers)}"

            cipher_suite = Fernet(KEY)
            cipher_text = cipher_suite.encrypt(credentials.encode('utf-8'))
            with open('dns.cfg', 'w') as f:
                f.write(cipher_text.decode('utf-8'))
            logs.set("INFO - DNS servers are saved to a file (dns.cfg).")

        close_dns_window()

    if not dns_servers:
        import_credentials("dns")

    if len(dns_servers) < 4:
        dns_servers.extend([""] * (4 - len(dns_servers)))

    dns_window = tk.Toplevel(gui)
    dns_window.title('DNS Servers')
    dns_window.focus_force()
    dns_window.grab_set()
    dns_window.focus_set()
    dns_window.transient(gui)
    gui.attributes('-disabled', True)

    dns_window.geometry('%dx%d+%d+%d' % (350, 230, gui.winfo_rootx() + 100, gui.winfo_rooty() + 70))
    dns_window.resizable(False, False)
    dns_window.iconphoto(False, icon)

    dns_frame = tk.Frame(dns_window)
    dns_frame.pack()

    dns_label = ttk.Label(dns_frame, text='DNS Servers (IPv4)', font=("Arial bold", 10))
    dns_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    dns_svr1_label = ttk.Label(dns_frame, text='Primary:', font=("Arial bold", 8))
    dns_svr1_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')

    dns_svr1_entry = ttk.Entry(dns_frame, width=40)
    dns_svr1_entry.bind("<Button-3>", right_click)
    dns_svr1_entry.insert(0, dns_servers[0])
    dns_svr1_entry.grid(row=1, column=1, padx=5, pady=5)

    dns_svr2_label = ttk.Label(dns_frame, text='Secondary:', font=("Arial", 8))
    dns_svr2_label.grid(row=2, column=0, padx=5, pady=5, sticky='e')

    dns_svr2_entry = ttk.Entry(dns_frame, width=40)
    dns_svr2_entry.bind("<Button-3>", right_click)
    dns_svr2_entry.insert(0, dns_servers[1])
    dns_svr2_entry.grid(row=2, column=1, padx=5, pady=5)

    dns_svr3_label = ttk.Label(dns_frame, text='Tertiary:', font=("Arial", 8))
    dns_svr3_label.grid(row=3, column=0, padx=5, pady=5, sticky='e')

    dns_svr3_entry = ttk.Entry(dns_frame, width=40)
    dns_svr3_entry.bind("<Button-3>", right_click)
    dns_svr3_entry.insert(0, dns_servers[2])
    dns_svr3_entry.grid(row=3, column=1, padx=5, pady=5)

    dns_svr4_label = ttk.Label(dns_frame, text='Quaternary:', font=("Arial", 8))
    dns_svr4_label.grid(row=4, column=0, padx=5, pady=5, sticky='e')

    dns_svr4_entry = ttk.Entry(dns_frame, width=40)
    dns_svr4_entry.bind("<Button-3>", right_click)
    dns_svr4_entry.insert(0, dns_servers[3])
    dns_svr4_entry.grid(row=4, column=1, padx=5, pady=5)

    remember_me = tk.IntVar()
    remember_me_checkbutton = ttk.Checkbutton(dns_frame, text='Remember Me?', variable=remember_me,
                                              onvalue=1, offvalue=0, cursor='hand2', takefocus=False, width=15)
    remember_me_checkbutton.grid(row=5, column=0, columnspan=2, padx=25, pady=5, sticky='w')

    buttons_frame = ttk.Frame(dns_frame)
    buttons_frame.grid(row=6, column=0, columnspan=2, padx=5, pady=5)

    save_button = ttk.Button(buttons_frame, text='Save', cursor='hand2', width=15,
                             command=lambda: save_dns_servers(remember_me.get()))
    save_button.grid(row=0, column=0, padx=5, pady=5, sticky='e')

    cancel_button = ttk.Button(buttons_frame, text='Cancel', cursor='hand2', width=15,
                               command=lambda: close_dns_window())
    cancel_button.grid(row=0, column=1, padx=5, pady=5, sticky='w')

    dns_window.protocol("WM_DELETE_WINDOW", lambda: close_dns_window())
    dns_window.bind('<Return>', lambda event: save_dns_servers(remember_me.get()))

    dns_window.mainloop()


################################################################

# Searching Functions
def get_subnet_type(subnet):
    global separators
    subnet = str(subnet).strip()

    if isinstance(subnet, str):
        subnet = subnet.strip()
        try:  # IP
            if subnet.endswith("/32") and ipaddress.ip_address(subnet[:-3]):
                return "IP"
            raise Exception
        except:
            try:  # IP
                if subnet.endswith("/128") and ipaddress.ip_address(subnet[:-4]):
                    return "IP"
                raise Exception
            except:
                try:  # IP
                    if "/" not in subnet and ipaddress.ip_address(subnet):
                        return "IP"
                    raise Exception
                except:
                    try:  # Subnet
                        if ipaddress.ip_network(subnet, strict=False):
                            return "Subnet"
                        raise Exception
                    except:
                        try:  # Range
                            if "-" in subnet:
                                subnet = subnet.split("-")
                                if len(subnet) == 2 and ipaddress.ip_address(subnet[0]) and ipaddress.ip_address(
                                        subnet[1]):
                                    return "Range"
                            raise Exception
                        except:
                            try:  # List
                                if any(sep in subnet for sep in separators):
                                    sep = [sep for sep in separators if sep in subnet][0]
                                    subnet = subnet.split(sep)
                                    for ip in subnet:
                                        if get_subnet_type(ip) == "Invalid":
                                            return "Invalid"
                                    return "List"
                                raise Exception
                            except:
                                return "Invalid"

    elif isinstance(subnet, ipaddress.IPv4Address) or isinstance(subnet, ipaddress.IPv6Address):
        return "IP"

    elif isinstance(subnet, ipaddress.IPv4Network) or isinstance(subnet, ipaddress.IPv6Network):
        return "Subnet"

    else:
        return "Invalid"


def check_ip_in_subnet(ip, ip_type, subnet, subnet_type):
    global separators, logs

    try:
        if isinstance(ip, str):
            ip = convert_address_object(ip, ip_type)

        if isinstance(subnet, str):
            subnet = convert_address_object(subnet, subnet_type)

        if ((ip_type == "Invalid" or subnet_type == "Invalid") or
                ((ip_type == "Subnet" or ip_type == "Range" or ip_type == "List") and subnet_type == "IP")):
            return False

        elif ip_type == "IP" and subnet_type == "IP":
            if ip == subnet:
                return True

        elif ip_type == "IP" and subnet_type == "Subnet":
            if ip in subnet:
                return True

        elif ip_type == "IP" and subnet_type == "Range":
            if isinstance(subnet, list):
                if subnet[0] <= ip <= subnet[1]:
                    return True

        elif ip_type == "IP" and subnet_type == "List":
            if isinstance(subnet, list):
                for subnet_tmp in subnet:
                    if check_ip_in_subnet(ip, ip_type, subnet_tmp, get_subnet_type(subnet_tmp)):
                        return True

        elif ip_type == "Subnet" and subnet_type == "Subnet":
            if ((ip == subnet) or
                    (
                            ip.network_address >= subnet.network_address and ip.broadcast_address <= subnet.broadcast_address)):
                return True

        elif ip_type == "Subnet" and subnet_type == "Range":
            if isinstance(ip, list):
                if ip.network_address >= subnet[0] and ip.broadcast_address <= subnet[1]:
                    return True

        elif ip_type == "Subnet" and subnet_type == "List":
            if isinstance(subnet, list):
                for subnet_tmp in subnet:
                    if check_ip_in_subnet(ip, ip_type, subnet_tmp, get_subnet_type(subnet_tmp)):
                        return True

        elif ip_type == "Range" and subnet_type == "Subnet":
            if isinstance(ip, list):
                if ip[0] >= subnet.network_address and ip[1] <= subnet.broadcast_address:
                    return True

        elif ip_type == "Range" and subnet_type == "Range":
            if isinstance(ip, list) and isinstance(subnet, list):
                if ip[0] >= subnet[0] and ip[1] <= subnet[1]:
                    return True

        elif ip_type == "Range" and subnet_type == "List":
            if isinstance(ip, list) and isinstance(subnet, list):
                # ip = f"{ip[0]}-{ip[1]}"
                for subnet_tmp in subnet:
                    if check_ip_in_subnet(ip, ip_type, subnet_tmp, get_subnet_type(subnet_tmp)):
                        return True

        elif ip_type == "List" and subnet_type == "Subnet":
            if isinstance(ip, list):
                flag = 0
                for ip_tmp in ip:
                    if check_ip_in_subnet(ip_tmp, get_subnet_type(ip_tmp), subnet, subnet_type):
                        flag += 1
                if flag == len(ip):
                    return True

        elif ip_type == "List" and subnet_type == "Range":
            if isinstance(ip, list) and isinstance(subnet, list):
                flag = 0
                for ip_tmp in ip:
                    if check_ip_in_subnet(ip_tmp,
                                          get_subnet_type(ip_tmp),
                                          subnet,
                                          subnet_type):
                        flag += 1
                if flag == len(ip):
                    return True

        elif ip_type == "List" and subnet_type == "List":
            if isinstance(ip, list) and isinstance(subnet, list):
                flag = 0
                for ip_tmp in ip:
                    for subnet_tmp in subnet:
                        if check_ip_in_subnet(ip_tmp,
                                              get_subnet_type(ip_tmp),
                                              subnet_tmp,
                                              get_subnet_type(subnet_tmp)):
                            flag += 1
                            break
                if flag == len(ip):
                    return True

        return False

    except Exception as e:
        if not isinstance(ip, str):
            ip = convert_address_object(ip, ip_type)
        if not isinstance(subnet, str):
            subnet = convert_address_object(subnet, subnet_type)

        logs.set(f"ERROR - Failed checking {ip} in {subnet}: {str(e)}")
        return False


def convert_address_object(address, address_type):
    global separators
    tmp = address

    try:
        if isinstance(address, str) and isinstance(address_type, str):
            address = address.strip()

            if address_type == "IP":
                address = re.sub("/32", "", address)
                address = re.sub("/128", "", address)
                return ipaddress.ip_address(address)

            elif address_type == "Subnet":
                return ipaddress.ip_network(address, strict=False)

            elif address_type == "Range":
                address = address.split("-")
                address = [x.strip() for x in address]
                if ipaddress.ip_address(address[0]) > ipaddress.ip_address(address[1]):
                    address[0], address[1] = address[1], address[0]
                return [ipaddress.ip_address(address[0]), ipaddress.ip_address(address[1])]

            elif address_type == "List":
                if any(sep in address for sep in separators):
                    sep = [sep for sep in separators if sep in address][0]
                    address = address.split(sep)
                    address = [x.strip() for x in address]
                    return [convert_address_object(ip, get_subnet_type(ip)) for ip in address]

            else:
                return address

        elif isinstance(address_type, str):
            if address_type == "IP":
                return str(address)

            elif address_type == "Subnet":
                return str(address)

            elif address_type == "Range":
                return f"{str(address[0])}-{str(address[1])}"

            elif address_type == "List":
                return "\n".join([str(ip) for ip in address])

            else:
                return address
    except:
        return tmp


def save_to_excel():
    global outputs, logs
    try:
        headers = ["Subnet", "Reference", "Tenant / Vsys", "Address Object", "Ref Subnet", "Type", "S/M"]
        filename = f"Results_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx"

        workbook = xlsxwriter.Workbook(filename)
        header_format = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})
        text_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter'})

        for sheet in outputs.keys():
            worksheet = workbook.add_worksheet(sheet)
            worksheet.write_row('A1', headers, header_format)
            for i in range(len(outputs[sheet][0])):
                worksheet.write_row(f'A{i + 2}', [outputs[sheet][x][i] for x in range(len(outputs[sheet]))],
                                    text_format)
            worksheet.autofit()

        workbook.close()
        logs.set(f"INFO - Results are saved into ({filename}) file")
        return filename
    except Exception as e:
        logs.set(f"ERROR - Failed saving results to Excel file: {str(e)}")
        return "False"


def start_begin():
    global start_flag, button_start, s
    start_flag = True
    s = ttk.Style()
    s.configure("start.TButton", font=("Arial bold", 15), foreground="red")
    button_start.config(text="Stop", state="normal", cursor="hand2", style="start.TButton", command=stop_process)
    disable_buttons_all()


def start_end():
    global start_flag, button_start, start_thread, s
    start_flag = False
    try:
        if start_thread.is_alive():
            start_thread.terminate()
    except:
        pass
    s = ttk.Style()
    s.configure("start.TButton", font=("Arial bold", 15), foreground="green")
    button_start.config(text="Start", state="normal", cursor="hand2", style="start.TButton", command=start)
    enable_buttons_all()


def start():
    global start_thread, logs, button_start, inputs, inputs_no, input_invalids, flags, ref_checkbox_var, refs, ref_no, \
        panorama_checkbox_var, panorama_addresses, panorama_ip, panorama_username, panorama_password, panorama_vsys, \
        forti_checkbox_var, forti_ip, forti_port, forti_username, forti_password, forti_vdom, dns_checkbox_var, resolvers, \
        dns_resolver, dns_servers, outputs, start_flag, gui, input_file, ref_file

    start_begin()
    button_start.config(command=lambda: False)
    logs.set("INFO - Starting the process...")

    # Check input file
    if not inputs:
        tmp_input = PropagatingThread(target=check_input_file, daemon=True)
        tmp_input.start()
        while tmp_input.is_alive():
            gui.update()
            time.sleep(0.1)
        tmp_input = tmp_input.join()
        if not tmp_input:
            set_info1("ERROR", "red", "arrow")
            set_info2("Input file is empty or not valid.", "red", "arrow")
            logs.set("ERROR - Input file is empty or not valid.")
            start_end()
            return False
    if inputs:
        if not inputs_no:
            set_info1("ERROR", "red", "arrow")
            set_info2("No valid records are found in the input file.", "red", "arrow")
            logs.set(f"ERROR - No valid records are found in the input file ({input_file}).")
            start_end()
            return False

        all_inputs_no = inputs_no + len(input_invalids)

        set_info1("INFO", "green", "arrow")
        set_info2(f"Input file is valid. {all_inputs_no} inputs are found.", "green", "arrow")
        logs.set(f"INFO - Input file is valid. {all_inputs_no} inputs are found.")

    # Check Used Methods
    check_methods()

    if flags[0]:
        # Check reference file
        if not refs:
            tmp_ref = PropagatingThread(target=check_ref_file, daemon=True)
            tmp_ref.start()
            while tmp_ref.is_alive():
                gui.update()
                time.sleep(0.1)
            tmp_ref = tmp_ref.join()
            if not tmp_ref:
                start_end()
                ref_checkbox_var.set(0)
                set_info1("ERROR", "red", "arrow")
                set_info2("Reference file is empty or not valid.", "red", "arrow")
                logs.set("ERROR - Reference file is empty or not valid.")
                return
        if refs:
            if not ref_no:
                start_end()
                ref_checkbox_var.set(0)
                set_info1("ERROR", "red", "arrow")
                set_info2("No valid records are found in the reference file.", "red", "arrow")
                logs.set(f"ERROR - No valid records are found in the reference file ({ref_file}).")
                return

            set_info1("INFO", "green", "arrow")
            set_info2(f"Reference file is valid. {str(ref_no)} references are found.", "green", "arrow")
            logs.set(f"INFO - Reference file is valid. {str(ref_no)} references are found.")

    if flags[1]:
        if not panorama_addresses:
            # Connect to Panorama
            tmp_connect_panorama = PropagatingThread(target=connect_to_panorama,
                                                     args=(
                                                         panorama_ip, panorama_username, panorama_password,
                                                         panorama_vsys),
                                                     daemon=True)
            tmp_connect_panorama.start()
            while tmp_connect_panorama.is_alive():
                gui.update()
                time.sleep(0.1)
            tmp_connect_panorama = tmp_connect_panorama.join()
            if not tmp_connect_panorama:
                start_end()
                panorama_checkbox_var.set(0)
                set_info1("ERROR", "red", "arrow")
                set_info2("Panorama is not connected.", "red", "arrow")
                logs.set("ERROR - Panorama is not connected.")
                return
            else:
                start_begin()
                tmp_export_panorama = PropagatingThread(target=export_panorama_address_objects, daemon=True)
                tmp_export_panorama.start()
                while tmp_export_panorama.is_alive():
                    gui.update()
                    time.sleep(0.1)
                tmp_export_panorama = tmp_export_panorama.join()
                if not tmp_export_panorama:
                    start_end()
                    panorama_checkbox_var.set(0)
                    set_info1("ERROR", "red", "arrow")
                    set_info2("No address objects are found in Panorama.", "red", "arrow")
                    logs.set("ERROR - No address objects are found in Panorama.")
                    return
                else:
                    set_info1("INFO", "green", "arrow")
                    set_info2(f"Panorama is connected. {str(len(panorama_addresses))} address objects are found.",
                              "green", "arrow")
                    logs.set(f"INFO - Panorama is connected. {str(len(panorama_addresses))} address objects are found.")
        else:
            set_info1("INFO", "green", "arrow")
            set_info2(f"Panorama is connected. {str(len(panorama_addresses))} address objects are found.",
                      "green", "arrow")
            logs.set(f"INFO - Panorama is connected. {str(len(panorama_addresses))} address objects are found.")

    if flags[3]:
        if not forti_addresses:
            # Connect to FortiGate
            tmp_connect_forti = PropagatingThread(target=connect_to_forti,
                                                  args=(
                                                      forti_ip, forti_port, forti_username, forti_password, forti_vdom),
                                                  daemon=True)
            tmp_connect_forti.start()
            while tmp_connect_forti.is_alive():
                gui.update()
                time.sleep(0.1)
            tmp_connect_forti = tmp_connect_forti.join()
            if not tmp_connect_forti:
                start_end()
                forti_checkbox_var.set(0)
                set_info1("ERROR", "red", "arrow")
                set_info2("FortiGate is not connected.", "red", "arrow")
                logs.set("ERROR - FortiGate is not connected.")
                return
            else:
                start_begin()
                tmp_export_forti = PropagatingThread(target=export_forti_address_objects, daemon=True)
                tmp_export_forti.start()
                while tmp_export_forti.is_alive():
                    gui.update()
                    time.sleep(0.1)
                tmp_export_forti = tmp_export_forti.join()
                if not tmp_export_forti:
                    start_end()
                    forti_checkbox_var.set(0)
                    set_info1("ERROR", "red", "arrow")
                    set_info2("No address objects are found in FortiGate.", "red", "arrow")
                    logs.set("ERROR - No address objects are found in FortiGate.")
                    return
                else:
                    set_info1("INFO", "green", "arrow")
                    set_info2(f"FortiGate is connected. {str(len(forti_addresses))} address objects are found.",
                              "green", "arrow")
                    logs.set(f"INFO - FortiGate is connected. {str(len(forti_addresses))} address objects are found.")
        else:
            set_info1("INFO", "green", "arrow")
            set_info2(f"FortiGate is connected. {str(len(forti_addresses))} address objects are found.",
                      "green", "arrow")
            logs.set(f"INFO - FortiGate is connected. {str(len(forti_addresses))} address objects are found.")

    if flags[2]:
        if not aci_addresses:
            # Connect to ACI
            tmp_connect_aci = PropagatingThread(target=connect_to_aci,
                                                args=(aci_ip, aci_username, aci_password, aci_class),
                                                daemon=True)
            tmp_connect_aci.start()
            while tmp_connect_aci.is_alive():
                gui.update()
                time.sleep(0.1)
            tmp_connect_aci = tmp_connect_aci.join()
            if not tmp_connect_aci:
                start_end()
                aci_checkbox_var.set(0)
                set_info1("ERROR", "red", "arrow")
                set_info2("ACI is not connected.", "red", "arrow")
                logs.set("ERROR - ACI is not connected.")
                return

            else:
                start_begin()
                tmp_export_aci = PropagatingThread(target=export_aci_address_objects, daemon=True)
                tmp_export_aci.start()
                while tmp_export_aci.is_alive():
                    gui.update()
                    time.sleep(0.1)
                tmp_export_aci = tmp_export_aci.join()
                if not tmp_export_aci:
                    start_end()
                    aci_checkbox_var.set(0)
                    set_info1("ERROR", "red", "arrow")
                    set_info2("No address objects are found in ACI.", "red", "arrow")
                    logs.set("ERROR - No address objects are found in ACI.")
                    return

                else:
                    set_info1("INFO", "green", "arrow")
                    set_info2(f"ACI is connected. {str(len(aci_addresses))} address objects are found.",
                              "green", "arrow")
                    logs.set(f"INFO - ACI is connected. {str(len(aci_addresses))} address objects are found.")
        else:
            set_info1("INFO", "green", "arrow")
            set_info2(f"ACI is connected. {str(len(aci_addresses))} address objects are found.", "green", "arrow")
            logs.set(f"INFO - ACI is connected. {str(len(aci_addresses))} address objects are found.")

    if flags[4]:
        if not resolvers:
            check_tmp = PropagatingThread(target=check_dns_servers,
                                          args=(dns_servers,),
                                          daemon=True)
            check_tmp.start()
            while check_tmp.is_alive():
                gui.update()
                time.sleep(0.1)
            resolvers = check_tmp.join()

            if not resolvers:
                start_end()
                dns_checkbox_var.set(0)
                set_info1("ERROR", "red", "arrow")
                set_info2("No reachable DNS servers are found.", "red", "arrow")
                logs.set("ERROR - No reachable DNS servers are found.")
                return
            else:
                dns_resolver = dns.resolver.Resolver()
                dns_resolver.nameservers = resolvers
                set_info1("INFO", "green", "arrow")
                set_info2(f"{len(resolvers)} DNS servers are reachable.", "green", "arrow")
                logs.set(f"INFO - {len(resolvers)} DNS servers are reachable.")
        else:
            set_info1("INFO", "green", "arrow")
            set_info2(f"{len(resolvers)} DNS servers are found.", "green", "arrow")
            logs.set(f"INFO - {len(resolvers)} DNS servers are found.")

    set_info1("INFO", "green", "arrow")
    set_info2("The process is ready to start.", "green", "arrow")
    button_start.config(command=stop_process)

    invalids_str = ""

    if input_invalids:
        invalids_str += f"Invalid Inputs: ({len(input_invalids)} records)\n\n"
        for sheet, sub, row in input_invalids:
            sub = re.sub("\n+", "[NL]", sub)
            # if len(sub) > 15:
            #     sub = f"{sub[:15]}..."
            invalids_str += f"Sheet: {sheet}\tRow: {row}\tSubnet: {sub}\n"
        invalids_str += "\n"

    if ref_invalids:
        if invalids_str:
            invalids_str += "=" * 50 + "\n\n"
        invalids_str += f"Invalid References: ({len(ref_invalids)} records)\n\n"
        for sheet, sub, row in ref_invalids:
            sub = re.sub("\n+", "[NL]", sub)
            # if len(sub) > 15:
            #     sub = f"{sub[:15]}..."
            invalids_str += f"Sheet: {sheet}\tRow: {row}\tSubnet: {sub}\n"

    if invalids_str:
        with open("invalids.txt", "w") as f:
            if "[NL]" in invalids_str:
                f.write("[NL] == New Line\n\n")
            f.write(invalids_str.strip())

    if len(inputs.keys()) == 1:
        sheet_str = f"{len(inputs.keys())} Sheet"
    else:
        sheet_str = f"{len(inputs.keys())} Sheets"

    content_str = f"Input File Contents:\n{sheet_str}, {inputs_no} Valid Subnets, {len(input_invalids)} Invalid Subnets."

    if input_invalids:
        content_str += "\nInvalid details are saved in the 'invalids.txt' file.\n\n"
    else:
        content_str += "\n\n"

    if flags[0]:
        if len(refs.keys()) == 1:
            sheet_str = f"{len(refs.keys())} Sheet"
        else:
            sheet_str = f"{len(refs.keys())} Sheets"

        content_str += f"Reference File Contents:\n{sheet_str}, {ref_no} Valid Subnets, {len(ref_invalids)} Invalid Subnets."

        if ref_invalids:
            content_str += "\nInvalid details are saved in the 'invalids.txt' file.\n\n"
        else:
            content_str += "\n\n"

    if flags[1]:
        content_str += f"Panorama Address Objects:\n{len(panorama_addresses)} Address Objects.\n\n"

    if flags[3]:
        content_str += f"FortiGate Address Objects:\n{len(forti_addresses)} Address Objects.\n\n"

    if flags[2]:
        content_str += f"ACI Address Objects:\n{len(aci_addresses)} Address Objects.\n\n"

    if flags[4]:
        content_str += f"DNS Servers:\n{len(resolvers)} reachable DNS Servers.\n\n"

    ans = tk.messagebox.askquestion("Info",
                                    f"{content_str}\nClick (Yes) to continue.",
                                    parent=gui)
    if ans != "yes":
        set_info1("INFO", "green", "arrow")
        set_info2("The operation has been canceled.", "red", "arrow")
        logs.set("INFO - The user canceled the searching.")
        start_end()
        return False

    start_thread = PropagatingThread(target=search)
    start_thread.start()


def stop_process():
    global gui, start_flag, logs, progress_label, progress_bar, stop_flag
    stop_flag = True

    ans = tk.messagebox.askyesnocancel("Warning",
                                       "Do you want to save the results before stopping the process?",
                                       icon="warning",
                                       parent=gui)

    if ans is True:
        logs.set("INFO - The user stopped the process and saving the results.")
        start_flag = False

    elif ans is False:
        logs.set("INFO - The user stopped the process without saving the results.")

        if progress_bar['value'] >= progress_bar['maximum']:
            progress_label.config(text="%100", foreground="green")
        else:
            progress_label.config(foreground="red")

        set_info1(f"INFO", "green", "arrow")
        set_info2(f"The process has been stopped.", "red", "arrow")

        start_end()
        clear_variables()
        gui.update()

    stop_flag = False


def search():
    global gui, inputs, refs, outputs, logs, progress_label, progress_bar, inputs_no, start_flag, stop_flag, \
        input_invalids, refs_bkp, inputs_bkp, panorama_addresses, panorama_addresses_bkp, aci_addresses, \
        aci_addresses_bkp, forti_addresses, forti_addresses_bkp, dns_resolver, dns_servers, flags

    start_begin()
    inputs_no = inputs_no + len(input_invalids)
    progress_bar.config(value=0, maximum=inputs_no)
    progress_label.config(text="%0", foreground="black")

    set_info1(f"INFO", "green", "arrow")
    logs.set("INFO - The search has been started.")

    count = 0
    count2 = 0
    t1 = time.time()
    for sheet in inputs.keys():
        count2 += 1
        outputs[sheet] = [[] for _ in range(7)]
        for i, subnet in enumerate(inputs[sheet]):

            while stop_flag:
                if not start_flag:
                    break
            else:
                if not start_flag:
                    break

            outputs[sheet][0].append(inputs_bkp[sheet][i][0])

            count += 1
            t2 = time.time()
            set_info2(
                f"Time Elapsed: {datetime.timedelta(seconds=round(t2 - t1))}  -  Time Remaining: {datetime.timedelta(seconds=round(((t2 - t1) / count) * (inputs_no - count)))}",
                "green", "arrow")
            progress_label.config(text=f"%{int((count / inputs_no) * 100)}")
            progress_bar.config(value=count)
            gui.update()

            if subnet[1] == "Invalid":
                outputs[sheet][1].append("")
                outputs[sheet][2].append("")
                outputs[sheet][3].append(inputs_bkp[sheet][i][0])
                outputs[sheet][4].append("")
                outputs[sheet][5].append("Invalid")
                outputs[sheet][6].append("None")

            else:
                ref_temp = []
                tenant_temp = []
                object_temp = []
                address_temp = []
                type_temp = []

                if flags[0]:
                    for ref_sheet in refs.keys():
                        for j, ref_subnet in enumerate(refs[ref_sheet]):
                            if check_ip_in_subnet(subnet[0], subnet[1], ref_subnet[2], ref_subnet[3]):
                                ref_temp.append(f"Ref File - {ref_sheet} row {j + 2}")
                                tenant_temp.append(ref_subnet[0])
                                object_temp.append(ref_subnet[1])
                                address_temp.append(refs_bkp[ref_sheet][j][2])
                                type_temp.append(ref_subnet[3])

                if flags[1]:
                    for j, ref_subnet in enumerate(panorama_addresses):
                        if check_ip_in_subnet(subnet[0], subnet[1], ref_subnet[2], ref_subnet[3]):
                            ref_temp.append(ref_subnet[0])
                            tenant_temp.append(ref_subnet[0][6:])
                            object_temp.append(ref_subnet[1])
                            address_temp.append(panorama_addresses_bkp[j][2])
                            type_temp.append(ref_subnet[3])

                if flags[3]:
                    for j, ref_subnet in enumerate(forti_addresses):
                        if check_ip_in_subnet(subnet[0], subnet[1], ref_subnet[2], ref_subnet[3]):
                            ref_temp.append(ref_subnet[0])
                            tenant_temp.append(ref_subnet[0][6:])
                            object_temp.append(ref_subnet[1])
                            address_temp.append(forti_addresses_bkp[j][2])
                            type_temp.append(ref_subnet[3])

                if flags[2]:
                    for j, ref_subnet in enumerate(aci_addresses):
                        if check_ip_in_subnet(subnet[0], subnet[1], ref_subnet[2], ref_subnet[3]):
                            ref_temp.append(ref_subnet[0])
                            tenant_temp.append(ref_subnet[0][6:])
                            object_temp.append(ref_subnet[1])
                            address_temp.append(aci_addresses_bkp[j][2])
                            type_temp.append(ref_subnet[3])

                if flags[4]:
                    if subnet[1] == "IP":
                        try:
                            ref_subnet = str(dns_resolver.resolve_address(subnet[0].compressed, search=True)[0])
                            ref_temp.append("DNS")
                            tenant_temp.append("DNS")
                            object_temp.append(ref_subnet)
                            address_temp.append(subnet[0].compressed)
                            type_temp.append("Domain Name")
                        except:
                            pass

                if len(ref_temp) > 1:
                    outputs[sheet][6].append("Multiple")
                elif len(ref_temp) == 1:
                    outputs[sheet][6].append("Single")

                if len(ref_temp) != 0:
                    outputs[sheet][1].append("\n".join(ref_temp))
                    outputs[sheet][2].append("\n".join(tenant_temp))
                    outputs[sheet][3].append("\n".join(object_temp))
                    outputs[sheet][4].append("\n".join(address_temp))
                    outputs[sheet][5].append("\n".join(type_temp))
                    continue

                outputs[sheet][1].append("")
                outputs[sheet][2].append("")
                outputs[sheet][3].append(inputs_bkp[sheet][i][0])
                outputs[sheet][4].append("")
                outputs[sheet][5].append("Not Found")
                outputs[sheet][6].append("None")
        else:
            if not start_flag:
                break

    if progress_bar['value'] >= progress_bar['maximum']:
        progress_label.config(text="%100", foreground="green")
    else:
        progress_label.config(foreground="red")

    t2 = time.time()
    filename = save_to_excel()

    if filename != "False":
        set_info1(f"INFO - Time Elapsed: {datetime.timedelta(seconds=round(t2 - t1))}", "green", "arrow")
        set_info2(f"Outputs are saved into '{filename}' file.", "green", "hand2")
    else:
        set_info1(f"ERORR - Time Elapsed: {datetime.timedelta(seconds=round(t2 - t1))}", "red", "arrow")
        set_info2(f"Failed saving outputs to Excel file.", "red", "arrow")
    info_label_2.bind("<Button-1>", lambda e: os.startfile(filename))
    logs.set(f"INFO - Time Elapsed: {datetime.timedelta(seconds=round(t2 - t1))}")

    start_end()
    clear_variables()
    gui.update()


################################################################

right_click_fn()

# Create a title label
title_label = ttk.Label(gui, text=NAME, font=("Arial bold", 20))
title_label.pack(pady=(15, 0), anchor="center")

buttons_frame = ttk.Frame(gui)
buttons_frame.pack(pady=(0, 10), padx=20, anchor="center")

github_icon = github_icon.zoom(1).subsample(18)
github_button = tk.Button(buttons_frame, text=f" {AUTHOR}", font=("Arial bold", 12),
                          image=github_icon, borderwidth=0, relief="sunken", compound="left",
                          cursor="hand2", command=lambda: webbrowser.open_new(GITHUB), anchor="center")
github_button.grid(row=0, column=0, padx=(5, 300), pady=(3, 5), sticky="w")

info_icon = info_icon.zoom(2).subsample(13)
info_button = tk.Button(buttons_frame, text="Help ", font=("Arial bold", 12),
                        image=info_icon, borderwidth=0, relief="sunken", compound="right",
                        cursor="hand2", command=lambda: info_window(), anchor="center")
info_button.grid(row=0, column=1, padx=(0, 5), pady=(3, 5), sticky="e")

# Create a separator
separator_2 = ttk.Separator(gui, orient="horizontal")
separator_2.pack(fill="x", pady=(0, 10), padx=20)

# Create an info labels
info_label_1 = ttk.Label(gui, text="INFO", font=("Arial bold", 10), foreground="green", cursor="arrow")
info_label_2 = ttk.Label(gui, text="You must enter input file / IP and at least one valid searching method",
                         font=("Arial bold", 10), foreground="green", cursor="arrow")
info_label_1.pack(pady=0, anchor="center", padx=20)
info_label_2.pack(pady=(0, 10), anchor="center", padx=20)

# Create a separator
separator_3 = ttk.Separator(gui, orient="horizontal")
separator_3.pack(fill="x", pady=(0, 10), padx=20)

# Create a label frame
label_frame = ttk.Frame(gui)
label_frame.pack(pady=(0, 10), padx=20, anchor="center")

# Create a label
input_label = ttk.Label(label_frame, text="Input File / IP")
input_label.grid(row=0, column=0, sticky="e", padx=(0, 5), pady=(0, 5))

# Create a text entry
input_StringVar = tk.StringVar()
# if the value of the StringVar is changed, call the check_input_file function
input_StringVar.trace_add("write", check_input_file)
input_entry = ttk.Entry(label_frame, width=50, textvariable=input_StringVar, name="input_entry")
input_entry.bind("<Button-3>", right_click)
input_entry.bind("<FocusOut>", lambda e: check_input_file())
input_entry.grid(row=0, column=1, sticky="w", padx=(0, 5), pady=(0, 5))
# input_entry.focus()

# Create a button
input_button = ttk.Button(label_frame, text="Browse", cursor="hand2", command=browse_input)
input_button.grid(row=0, column=2, sticky="w", pady=(0, 5))

# Checkbox for the Input File
in_checkbox_var = tk.IntVar()
in_checkbox = ttk.Checkbutton(label_frame, variable=in_checkbox_var, onvalue=1, offvalue=0, takefocus=False,
                              command=check_methods)
in_checkbox_var.set(1)
in_checkbox.grid(row=0, column=3, sticky="w", pady=(0, 5), padx=5)
in_checkbox.config(state="disabled", cursor="arrow")

# Create a label
ref_label = ttk.Label(label_frame, text="Reference File")
ref_label.grid(row=1, column=0, sticky="e", padx=(0, 5), pady=(0, 5))

# Create a text entry
ref_StringVar = tk.StringVar()
# if the value of the StringVar is changed, call the check_ref_file function
ref_StringVar.trace_add("write", check_ref_file)
ref_entry = ttk.Entry(label_frame, width=50, textvariable=ref_StringVar, name="ref_entry")
ref_entry.bind("<Button-3>", right_click)
ref_entry.bind("<FocusOut>", lambda e: check_ref_file())
ref_entry.grid(row=1, column=1, sticky="w", padx=(0, 5), pady=(0, 5))

# Create a button
ref_button = ttk.Button(label_frame, text="Browse", cursor="hand2", command=browse_ref)
ref_button.grid(row=1, column=2, sticky="w", pady=(0, 5))

# Checkbox for the Reference File
ref_checkbox_var = tk.IntVar()
ref_checkbox = ttk.Checkbutton(label_frame, variable=ref_checkbox_var, onvalue=1, offvalue=0, takefocus=False,
                               command=check_methods)
# ref_checkbox_var.set(1)
ref_checkbox_var.trace_add("write", check_methods)
ref_checkbox.grid(row=1, column=3, sticky="w", pady=(0, 5), padx=5)

# Create a label
label_panorama = ttk.Label(label_frame, text="PAN Status")
label_panorama.grid(row=2, column=0, sticky="e", padx=(0, 5))

# Create a label
label_panorama_status = ttk.Label(label_frame, text="Not Connected", foreground="red", font=("Arial bold", 10))
label_panorama_status.grid(row=2, column=1, sticky="w", padx=(0, 5))

# Create a button
button_panorama = ttk.Button(label_frame, text="Connect", cursor="hand2", command=get_panorama_credentials)
# button_panorama.config(state="disabled", cursor="arrow")
button_panorama.grid(row=2, column=2, sticky="w")

# Checkbox for the Panorama
panorama_checkbox_var = tk.IntVar()
panorama_checkbox = ttk.Checkbutton(label_frame, variable=panorama_checkbox_var, onvalue=1, offvalue=0,
                                    takefocus=False,
                                    command=check_methods)
panorama_checkbox_var.set(0)
panorama_checkbox_var.trace_add("write", check_methods)
panorama_checkbox.grid(row=2, column=3, sticky="w", pady=(0, 5), padx=5)

# Create a label
label_forti = ttk.Label(label_frame, text="Forti Status")
label_forti.grid(row=3, column=0, sticky="e", padx=(0, 5))

# Create a label
label_forti_status = ttk.Label(label_frame, text="Not Connected", foreground="red", font=("Arial bold", 10))
label_forti_status.grid(row=3, column=1, sticky="w", padx=(0, 5))

# Create a button
button_forti = ttk.Button(label_frame, text="Connect", cursor="hand2", command=get_forti_credentials)
# button_forti.config(state="disabled", cursor="arrow")
button_forti.grid(row=3, column=2, sticky="w")

# Checkbox for the Forti
forti_checkbox_var = tk.IntVar()
forti_checkbox = ttk.Checkbutton(label_frame, variable=forti_checkbox_var, onvalue=1, offvalue=0, takefocus=False,
                                 command=check_methods)
forti_checkbox_var.set(0)
forti_checkbox_var.trace_add("write", check_methods)
forti_checkbox.grid(row=3, column=3, sticky="w", pady=(0, 5), padx=5)

# Create a label
label_aci = ttk.Label(label_frame, text="ACI Status")
label_aci.grid(row=4, column=0, sticky="e", padx=(0, 5))

# Create a label
label_aci_status = ttk.Label(label_frame, text="Not Connected", foreground="red", font=("Arial bold", 10))
label_aci_status.grid(row=4, column=1, sticky="w", padx=(0, 5))

# Create a button
button_aci = ttk.Button(label_frame, text="Connect", cursor="hand2", command=get_aci_credentials)
# button_aci.config(state="disabled", cursor="arrow")
button_aci.grid(row=4, column=2, sticky="w")

# Checkbox for the ACI
aci_checkbox_var = tk.IntVar()
aci_checkbox = ttk.Checkbutton(label_frame, variable=aci_checkbox_var, onvalue=1, offvalue=0, takefocus=False,
                               command=check_methods)
aci_checkbox_var.set(0)
aci_checkbox_var.trace_add("write", check_methods)
aci_checkbox.grid(row=4, column=3, sticky="w", pady=(0, 5), padx=5)

# Create a label
label_dns = ttk.Label(label_frame, text="DNS Status")
label_dns.grid(row=5, column=0, sticky="e", padx=(0, 5))

# Create a label
label_dns_status = ttk.Label(label_frame, text="Not Reachable", foreground="red", font=("Arial bold", 10))
label_dns_status.grid(row=5, column=1, sticky="w", padx=(0, 5))

# Create a button
button_dns = ttk.Button(label_frame, text="Check", cursor="hand2", command=get_dns_servers)
button_dns.grid(row=5, column=2, sticky="w")

# Checkbox for the DNS
dns_checkbox_var = tk.IntVar()
dns_checkbox = ttk.Checkbutton(label_frame, variable=dns_checkbox_var, onvalue=1, offvalue=0, takefocus=False,
                               command=check_methods)
dns_checkbox_var.set(0)
dns_checkbox_var.trace_add("write", check_methods)
dns_checkbox.grid(row=5, column=3, sticky="w", pady=(0, 5), padx=5)

# Create a separator
separator_4 = ttk.Separator(gui, orient="horizontal")
separator_4.pack(fill="x", pady=(0, 10), padx=20)

# Create a progress bar frame
progress_frame = ttk.Frame(gui)
progress_frame.pack(pady=(0, 10), padx=20, anchor="center")

# Create a progress bar
progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=450, mode="determinate")
progress_bar.grid(row=0, column=0, sticky="w")

# Create a label
progress_label = ttk.Label(progress_frame, text="%0", font=("Arial bold", 10))
progress_label.grid(row=0, column=1, padx=(10, 0), sticky="w")

# Create a separator
separator_5 = ttk.Separator(gui, orient="horizontal")
separator_5.pack(fill="x", pady=(0, 10), padx=20)

# Create a button
s = ttk.Style()
s.configure("start.TButton", font=("Arial bold", 15), foreground="green")
button_start = ttk.Button(gui, text="Start", cursor="hand2", style="start.TButton", command=start)
button_start.pack(pady=(0, 10), anchor="center", ipady=10, ipadx=75)

# GUI Bindings
gui.bind('<Escape>', close_gui_window)
gui.protocol("WM_DELETE_WINDOW", close_gui_window)

gui.iconphoto(False, icon)

# Import credential files
credentials = import_credentials()
if credentials[0]:
    PropagatingThread(target=connect_to_panorama,
                      args=(panorama_ip, panorama_username, panorama_password, panorama_vsys)).start()
if credentials[1]:
    PropagatingThread(target=connect_to_aci, args=(aci_ip, aci_username, aci_password, aci_class)).start()
if credentials[2]:
    PropagatingThread(target=connect_to_forti,
                      args=(forti_ip, forti_port, forti_username, forti_password, forti_vdom)).start()
PropagatingThread(target=check_dns_servers, args=dns_servers).start()

check_methods()

# Close the splash screen
try:
    if getattr(sys, 'frozen', False):
        pyi_splash.close()
except:
    pass


def OnFocusIn(event):
    global gui

    if type(event.widget).__name__ == 'Tk':
        event.widget.attributes('-topmost', False)

    gui.unbind('<FocusIn>', foucus_id)


gui.attributes('-topmost', True)
gui.focus_force()
foucus_id = gui.bind('<FocusIn>', OnFocusIn)

# Start the main loop
gui.mainloop()

logs.set("INFO - Application closed.")
logs.set("")
