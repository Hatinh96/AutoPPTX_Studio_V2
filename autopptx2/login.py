"""Màn hình đăng nhập Golden Asia — Supabase Auth."""
import threading
import tkinter as tk
import customtkinter as ctk

from .constants import ACCENT, ACCENT_HOVER, CARD, TEXT, MUTED, INPUT, BORDER, BG
from .constants import GA_BLUE_DEEP, GA_BLUE_GLOW, find_asset, LOGIN_LOGOS
from . import cloud as C


def build_login(host, app):
    """Vẽ form đăng nhập vào `host`. Gọi `app._try_login(email, pw, remember)`."""
    for w in host.winfo_children():
        w.destroy()
    host.configure(fg_color=BG)

    wrap = ctk.CTkFrame(host, fg_color='transparent')
    wrap.place(relx=0.5, rely=0.5, anchor='center')

    card = ctk.CTkFrame(wrap, corner_radius=20, width=860, height=520,
                        fg_color=CARD, border_width=1, border_color=BORDER)
    card.pack()
    card.pack_propagate(False)

    try:
        S = float(ctk.ScalingTracker.get_widget_scaling(host))
    except Exception:
        S = 1.0
    BW, BH = round(320 * S), round(520 * S)
    left = tk.Canvas(card, width=BW, height=BH, highlightthickness=0, bd=0,
                     bg=GA_BLUE_DEEP)
    left.pack(side='left', fill='y')
    pad = round(34 * S)
    y = round(180 * S)
    logo_drawn = False
    try:
        from PIL import Image, ImageTk
        logo_file = find_asset(*LOGIN_LOGOS)
        if logo_file:
            im = Image.open(logo_file).convert('RGBA')
            target_w = round(220 * S)
            target_h = max(1, int(im.height * target_w / im.width))
            im = im.resize((target_w, target_h), Image.Resampling.LANCZOS)
            photo_logo = ImageTk.PhotoImage(im)
            left.create_image(pad, round(90 * S), image=photo_logo, anchor='nw')
            left._logo_photo = photo_logo
            app._login_logo_photo = photo_logo
            y = round(90 * S) + target_h + round(24 * S)
            logo_drawn = True
    except Exception:
        logo_drawn = False
    if not logo_drawn:
        left.create_text(pad, y, text='GOLDEN ASIA', anchor='nw', fill='#ffffff',
                         font=('Segoe UI', 20, 'bold'))
        y += round(40 * S)

    left.create_text(pad, y, text='D I S C O V E R   T H E   D I F F E R E N C E',
                     anchor='nw', fill=GA_BLUE_GLOW, font=('Segoe UI', 8))
    y += round(36 * S)
    left.create_line(pad, y, pad + round(44 * S), y, fill='#ffffff', width=2)
    left.create_text(pad, BH - round(36 * S), text='AutoPPTX Studio V2',
                     anchor='nw', fill='#7fa8d8', font=('Segoe UI', 9))

    outer = ctk.CTkFrame(card, fg_color=CARD)
    outer.pack(side='left', fill='both', expand=True)
    ctk.CTkLabel(outer, text='© Golden Asia Media', font=ctk.CTkFont(size=11),
                 text_color=MUTED).pack(side='bottom', pady=(0, 16))

    form = ctk.CTkFrame(outer, fg_color=CARD)
    form.pack(expand=True, fill='x', padx=44)

    ctk.CTkLabel(form, text='Đăng nhập', font=ctk.CTkFont(size=25, weight='bold'),
                 text_color=TEXT, anchor='w').pack(fill='x')
    ctk.CTkLabel(form, text='Dùng tài khoản công ty (Supabase) để lấy Excel, avatar và ảnh nền trên đám mây.',
                 font=ctk.CTkFont(size=13), text_color=MUTED, anchor='w',
                 wraplength=380, justify='left').pack(fill='x', pady=(4, 20))

    def _lbl(t):
        ctk.CTkLabel(form, text=t, font=ctk.CTkFont(size=12, weight='bold'),
                     text_color=MUTED, anchor='w').pack(fill='x', pady=(0, 5))

    _lbl('Email')
    ent_email = ctk.CTkEntry(form, placeholder_text='ban@goldenasia.vn', height=44,
                             corner_radius=10, fg_color=INPUT, border_width=1,
                             border_color=BORDER, text_color=TEXT)
    ent_email.pack(fill='x', pady=(0, 14))

    _lbl('Mật khẩu')
    pw_row = ctk.CTkFrame(form, fg_color=CARD)
    pw_row.pack(fill='x')
    ent_pw = ctk.CTkEntry(pw_row, placeholder_text='••••••••', show='•', height=44,
                          corner_radius=10, fg_color=INPUT, border_width=1,
                          border_color=BORDER, text_color=TEXT)
    ent_pw.pack(side='left', fill='x', expand=True)
    shown = {'v': False}

    def _toggle():
        shown['v'] = not shown['v']
        ent_pw.configure(show='' if shown['v'] else '•')
        btn_eye.configure(text='🙈' if shown['v'] else '👁')

    btn_eye = ctk.CTkButton(pw_row, text='👁', width=46, height=44, corner_radius=10,
                            fg_color=INPUT, text_color=MUTED, hover_color=BORDER,
                            command=_toggle)
    btn_eye.pack(side='left', padx=(8, 0))

    for ent in (ent_email, ent_pw):
        ent.bind('<FocusIn>', lambda e, w=ent: w.configure(border_color=ACCENT))
        ent.bind('<FocusOut>', lambda e, w=ent: w.configure(border_color=BORDER))

    lbl_caps = ctk.CTkLabel(form, text='', font=ctk.CTkFont(size=11),
                            text_color='#e0a33e', anchor='w')
    lbl_caps.pack(fill='x', pady=(4, 0))

    def _caps(event):
        try:
            on = bool(event.state & 0x0002)
            lbl_caps.configure(text='⚠  Caps Lock đang bật' if on else '')
        except Exception:
            pass
    ent_pw.bind('<KeyPress>', _caps)

    remember = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(form, text='Nhớ mật khẩu trên máy này', variable=remember,
                    text_color=TEXT, font=ctk.CTkFont(size=13),
                    fg_color=ACCENT, hover_color=ACCENT_HOVER,
                    checkbox_width=19, checkbox_height=19,
                    corner_radius=5).pack(anchor='w', pady=(12, 2))

    lbl_err = ctk.CTkLabel(form, text='', font=ctk.CTkFont(size=12),
                           text_color='#e5534b', anchor='w', wraplength=380,
                           justify='left')

    btn = ctk.CTkButton(form, text='Đăng nhập', height=46, corner_radius=11,
                        fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color='#ffffff',
                        font=ctk.CTkFont(size=15, weight='bold'))
    btn.pack(fill='x', pady=(16, 0))

    widgets = {'email': ent_email, 'pw': ent_pw, 'eye': btn_eye,
               'btn': btn, 'err': lbl_err, 'remember': remember}

    def _go(_e=None):
        app._try_login(ent_email.get().strip(), ent_pw.get(),
                       bool(remember.get()), widgets)

    btn.configure(command=_go)
    for w in (ent_email, ent_pw):
        w.bind('<Return>', _go)

    host.after(120, ent_email.focus_set)

    def _apply_saved(saved):
        if not saved:
            return
        try:
            if not host.winfo_exists():
                return
            ent_email.insert(0, saved.get('email', ''))
            ent_pw.insert(0, saved.get('pw', ''))
            remember.set(True)
            ent_pw.focus_set()
        except Exception:
            pass

    def _load_saved_async():
        try:
            saved = C.load_saved_creds()
        except Exception:
            saved = None
        try:
            host.after(0, lambda: _apply_saved(saved))
        except Exception:
            pass

    threading.Thread(target=_load_saved_async, daemon=True).start()

    return widgets
