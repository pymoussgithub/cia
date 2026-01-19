import os
import time
import threading
import json
import hashlib
import tkinter as tk
from ttkbootstrap import Style
import customtkinter as ctk
from customtkinter import CTkImage
from tkinter import messagebox
from PIL import Image
from fenetre_principale import open_main_window




# ------------------ CONFIG ------------------

LOGO_PATH = "logo.png"
SCALE_LOGO = 0.7  # facteur de réduction du logo
WINDOW_WIDTH = 500
WINDOW_HEIGHT = 500
REMEMBER_FILE = "remember.txt"     # fichier pour mémoriser le username
USERS_FILE = "users.json"          # fichier stockant les users (username -> hashed_password)

# Admin password hash (SHA-256) — remplace par le hash correspondant à TON mot de passe admin si souhaité.
# Exemple ci-dessous correspond au mot de passe "admin123"
ADMIN_HASH = "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9"

# Par défaut, si users.json n'existe pas, on initialise avec 2 comptes d'exemple :
DEFAULT_USERS = {
    "admin": hashlib.sha256("1234".encode()).hexdigest(),
    "user": hashlib.sha256("abcd".encode()).hexdigest()
}

# ------------------ UTILITAIRES ------------------

def hash_password(password: str) -> str:
    """Retourne le hash sha256 hexadécimal du mot de passe"""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def load_users() -> dict:
    """Charge le dictionnaire username->hashed_password depuis users.json"""
    if not os.path.exists(USERS_FILE):
        # écrire les users par défaut
        try:
            with open(USERS_FILE, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_USERS, f, indent=2, ensure_ascii=False)
        except Exception:
            pass
        return DEFAULT_USERS.copy()
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return DEFAULT_USERS.copy()

def save_users(users: dict):
    """Sauvegarde le dictionnaire users dans users.json"""
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2, ensure_ascii=False)

# ------------------ SPLASH SCREEN ------------------

def show_splash():
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.configure(bg="white")

    if os.path.exists(LOGO_PATH):
        try:
            img = Image.open(LOGO_PATH)
            new_width = int(img.width * SCALE_LOGO)
            new_height = int(img.height * SCALE_LOGO)
            img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            logo = CTkImage(img_resized, size=(new_width, new_height))
            label = ctk.CTkLabel(splash, image=logo, text="")
            label.pack(pady=10)
        except Exception:
            pass

    splash.update_idletasks()
    w = splash.winfo_width()
    h = splash.winfo_height()
    x = (splash.winfo_screenwidth() // 2) - (w // 2)
    y = (splash.winfo_screenheight() // 2) - (h // 2)
    splash.geometry(f"{w}x{h}+{x}+{y}")

    time.sleep(2)
    splash.destroy()

# ------------------ APP PRINCIPALE ------------------

class LoginApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.configure(fg_color="white")
        self.title("Login")
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.resizable(False, False)
        self.attributes("-alpha", 0.0)
        self.lift()
        self.attributes("-topmost", True)
        self.after(200, lambda: self.attributes("-topmost", False))

        self.style = Style("flatly")
        ctk.set_appearance_mode("light")

        # chargement des users
        self.users = load_users()

        # Menu barre -> "Compte" -> "Créer un compte"
        self.create_menu()

        # Frame principal
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(anchor="n", fill="x", pady=(10,0))

        # --- LOGO ---
        if os.path.exists(LOGO_PATH):
            try:
                img = Image.open(LOGO_PATH)
                new_width = int(img.width * SCALE_LOGO)
                new_height = int(img.height * SCALE_LOGO)
                img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                self.logo = CTkImage(img_resized, size=(new_width, new_height))
                self.logo_label = ctk.CTkLabel(self.main_frame, image=self.logo, text="")
                self.logo_label.pack(pady=(0,10))
            except Exception:
                pass

        # --- Login form ---
        self.username_label = ctk.CTkLabel(self.main_frame, text="Nom d'utilisateur :", font=("Arial", 16))
        self.username_label.pack(pady=(5,2))
        self.username_entry = ctk.CTkEntry(self.main_frame, width=250)
        self.username_entry.pack(pady=(0,10))

        self.password_label = ctk.CTkLabel(self.main_frame, text="Mot de passe :", font=("Arial", 16))
        self.password_label.pack(pady=(5,2))
        self.password_entry = ctk.CTkEntry(self.main_frame, width=250, show="*")
        self.password_entry.pack(pady=(0,10))
        # Après avoir créé le password_entry
        self.password_entry.bind("<Return>", lambda event: self.check_login())

        # Case à cocher "Se souvenir du nom d'utilisateur"
        self.remember_var = tk.BooleanVar()
        self.remember_checkbox = ctk.CTkCheckBox(
            self.main_frame, text="Se souvenir du nom d'utilisateur",width=30, variable=self.remember_var, font=("Arial", 10)
        )
        self.remember_checkbox.pack(pady=(0,15))

        # Charger le username mémorisé si existe
        if os.path.exists(REMEMBER_FILE):
            try:
                with open(REMEMBER_FILE, "r", encoding="utf-8") as f:
                    saved_user = f.read().strip()
                    if saved_user:
                        self.username_entry.insert(0, saved_user)
                        self.remember_var.set(True)
            except:
                pass

        # Bouton Connexion avec Glow
        self.login_button = self.create_glow_button(self.main_frame, "Connexion", self.check_login)
        self.login_button.pack(pady=(15,10))

        self.fade_in()

    # ---- Menu creation ----
    def create_menu(self):
        menubar = tk.Menu(self)
        account_menu = tk.Menu(menubar, tearoff=0)
        account_menu.add_command(label="Créer un compte", command=self.open_create_account_dialog)
        menubar.add_cascade(label="Comptes", menu=account_menu)
        # attache la menu bar à la fenêtre principale (works with CTk since subclass of Tk)
        self.config(menu=menubar)

    # ---- Dialog pour création de compte ----
    def open_create_account_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Créer un compte")
        dlg.geometry("420x500")
        dlg.resizable(True, True)
        dlg.transient(self)
        dlg.grab_set()  # modal

        frame = ctk.CTkFrame(dlg, fg_color="transparent")
        frame.pack(expand=True, fill="both", padx=12, pady=12)

        tk.Label(frame, text="Nom d'utilisateur :").pack(anchor="w", pady=(6,2))
        entry_user = ctk.CTkEntry(frame, width=340)
        entry_user.pack(pady=(0,8))

        tk.Label(frame, text="Mot de passe :").pack(anchor="w", pady=(6,2))
        entry_pass = ctk.CTkEntry(frame, width=340, show="*")
        entry_pass.pack(pady=(0,8))

        tk.Label(frame, text="Confirmer mot de passe :").pack(anchor="w", pady=(6,2))
        entry_confirm = ctk.CTkEntry(frame, width=340, show="*")
        entry_confirm.pack(pady=(0,8))

        tk.Label(frame, text="Mot de passe admin :").pack(anchor="w", pady=(6,2))
        entry_admin = ctk.CTkEntry(frame, width=340, show="*")
        entry_admin.pack(pady=(0,10))

        # zone de message d'erreur
        err_label = ctk.CTkLabel(frame, text="", text_color="red")
        err_label.pack(pady=(0,6))

        def on_create():
            username = entry_user.get().strip()
            pwd = entry_pass.get()
            pwd2 = entry_confirm.get()
            admin_pwd = entry_admin.get()

            if not username:
                err_label.configure(text="Le nom d'utilisateur est requis.")
                return
            if ":" in username:
                err_label.configure(text="Caractère ':' non autorisé dans le nom d'utilisateur.")
                return
            if not pwd or not pwd2:
                err_label.configure(text="Les deux champs mot de passe sont requis.")
                return
            if pwd != pwd2:
                err_label.configure(text="Les mots de passe ne correspondent pas.")
                return
            if not admin_pwd:
                err_label.configure(text="Le mot de passe admin est requis.")
                return

            # vérifier la validité du mot de passe admin via le hash
            if hash_password(admin_pwd) != ADMIN_HASH:
                err_label.configure(text="Mot de passe admin incorrect.")
                return

            # vérifier si username déjà utilisé
            if username in self.users:
                err_label.configure(text="Nom d'utilisateur déjà existant.")
                return

            # ajouter l'utilisateur (hash du mot de passe)
            self.users[username] = hash_password(pwd)
            try:
                save_users(self.users)
            except Exception as e:
                err_label.configure(text=f"Erreur en sauvegarde: {e}")
                return

            messagebox.showinfo("Succès", f"Compte '{username}' créé avec succès.")
            dlg.destroy()

        btn_frame = tk.Frame(frame, bg="")
        btn_frame.pack(fill="x", pady=(6,0))
        ok_btn = ctk.CTkButton(btn_frame, text="Créer", width=120, command=on_create)
        ok_btn.pack(side="right", padx=6)
        cancel_btn = ctk.CTkButton(btn_frame, text="Annuler", width=120, command=dlg.destroy)
        cancel_btn.pack(side="right")

    # ---- Fade In ----
    def fade_in(self):
        for i in range(0, 101, 5):
            self.attributes("-alpha", i / 100)
            self.update()
            time.sleep(0.01)

    # ---- Glow Button ----
    def create_glow_button(self, parent, text, command):
        btn = ctk.CTkButton(
            parent,
            text=text,
            width=250,
            height=45,
            fg_color="#89B8E3",  # couleur du fond
            hover_color="#A1C9F1",  # couleur au survol
            text_color="white",  # couleur du texte
            font=("Arial", 16, "bold"),  # taille et style du texte
            command=command,
            corner_radius=10
        )

        # Glow / halo lumineux
        def on_enter(_):
            btn.configure(border_width=3, border_color="white")

        def on_leave(_):
            btn.configure(border_width=0)

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

        return btn

    # ---- Vérification login ----
    def check_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get()
        hashed = hash_password(password)
        remember = self.remember_var.get()

        if username in self.users and self.users[username] == hashed:
            if remember:
                with open(REMEMBER_FILE, "w", encoding="utf-8") as f:
                    f.write(username)
            else:
                if os.path.exists(REMEMBER_FILE):
                    os.remove(REMEMBER_FILE)

            # Récupérer dimensions écran avant destruction
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()

            self.destroy()  # fermer login

            # Appel de la fenêtre principale depuis main_window.py
            open_main_window(username, screen_width, screen_height)
        else:
            messagebox.showerror("Erreur", "Nom d'utilisateur ou mot de passe incorrect")


# ------------------ MAIN ------------------
if __name__ == "__main__":
    t = threading.Thread(target=show_splash)
    t.start()
    t.join()

    app = LoginApp()
    app.mainloop()
