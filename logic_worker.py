# logic_worker.py
# helper functions to safely interact with Tkinter from worker threads

import time
from tkinter import simpledialog

def thread_safe_askstring(root, title, prompt):
    """
    Safely call simpledialog.askstring from a background thread by scheduling on the mainloop
    and waiting for the result.
    """
    holder = {}

    def ask():
        holder["val"] = simpledialog.askstring(title, prompt)

    # schedule the dialog on the main thread
    root.after(0, ask)

    # wait for user to respond (short sleep to avoid busy spin)
    while "val" not in holder:
        time.sleep(0.05)

    return holder.get("val")

def thread_safe_update_label(root, label, text):
    """
    Safely update a label's text from a background thread.
    """
    root.after(0, lambda: label.config(text=text))
