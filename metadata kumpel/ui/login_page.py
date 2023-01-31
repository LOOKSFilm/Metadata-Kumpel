import customtkinter
from EditShareAPI import EsAuth
import threading

from ui.mainmenu_page import mainmenupage
from ui.widgets import LoadingAnimation




def loginpage(app, font, bg_color):
    def login(event):
        #frames = [".", "..", "...", "...."]
        frames = [" ", " ", " ", " "]
        loading_animation = LoadingAnimation(frame_status,frames,50,font=font,bg_color=bg_color)
        label_status.grid_forget()
        loading_animation.grid()
        label_status.configure(text="")
        global username
        username = input_username.get()
        password = input_password.get()
        
        connect = EsAuth.login("192.168.0.220", username, password)
        if connect == 200:
            app.unbind("<Return>")
            frame_login.grid_forget()
            mainmenupage(app, font, bg_color)
        else:
            label_status.configure(text="Wrong Username or Password")
            loading_animation.grid_forget()
            label_status.grid()
        
    def startlogin(event=True):
        threading.Thread(target=login, args=[event]).start()

    app.title("Metadata Kumpel: Login")
    frame_login = customtkinter.CTkFrame(app, fg_color=bg_color)
    frame_login.grid(sticky="NSEW")
    for i in range(3):
        frame_login.grid_rowconfigure(i, weight=1)
    frame_login.grid_columnconfigure(0, weight=1)
    input_username = customtkinter.CTkEntry(frame_login, placeholder_text="Username", justify="center", font=font)
    input_username.grid()
    input_password = customtkinter.CTkEntry(frame_login, placeholder_text="Password", show="*", justify="center", font=font)
    input_password.grid(sticky="N")
    btn_login = customtkinter.CTkButton(frame_login, text="Login", command=startlogin, font=font)
    app.bind("<Return>", startlogin)
    btn_login.grid(sticky="N")

    frame_status = customtkinter.CTkFrame(frame_login, bg_color="gray10")
    frame_status.grid(sticky="N", pady=17)
    label_status = customtkinter.CTkLabel(frame_status, text="", font=font, fg_color=bg_color)
    label_status.grid()
    print(btn_login.cget("fg_color"))


