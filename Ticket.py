import os
import time
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
from PIL import Image, ImageTk
import subprocess
from tkinter import font as tkfont
import random
from datetime import datetime, date, timedelta
from tkinter import ttk
import random   
import math

# Configurações de estilo
MODERN_FONT = "Segoe UI"
ACCENT_COLOR = "#4a6da7"
DARK_BG = "#121212"
LIGHT_BG = "#f5f7fa"
DARK_TEXT = "#e0e0e0"
LIGHT_TEXT = "#333333"

# Criar a janela principal
root = tk.Tk()
root.title("Análise de Tickets - DECSIS")
root.state("zoomed")  # fullscreen
root.resizable(True, True)

# Definir o ícone da janela
try:
    root.iconbitmap("C:/GOC/GOC_Launcher/Img/Icon/ds.ico")
except:
    pass

# Variáveis globais
dark_mode = False
logo_img = None
animation_jobs = []

# Configurar fonte moderna
try:
    custom_font = tkfont.Font(family=MODERN_FONT, size=12)
    root.option_add("*Font", custom_font)
except:
    pass

# Frame de fundo com gradiente
bg_canvas = tk.Canvas(root, highlightthickness=0)
bg_canvas.pack(fill="both", expand=True)

# Carregar imagens
def load_image(path, size):
    try:
        img = Image.open(path)
        img = img.resize(size, Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception as e:
        print(f"Erro ao carregar imagem: {e}")
        return None

# Animação de partículas
class Particle:
    def __init__(self, canvas):
        self.canvas = canvas
        self.size = random.randint(1, 2)
        x = random.randint(0, self.canvas.winfo_width())
        y = random.randint(0, self.canvas.winfo_height())
        self.dx = random.choice([-1, -0.5, 0.5, 1])  # Valores menores para velocidade horizontal
        self.dy = random.choice([-1, -0.5, 0.5, 1])  # Valores menores para velocidade vertical
        self.color = "black" if not dark_mode else "white"
        self.id = self.canvas.create_oval(
            x - self.size, y - self.size,
            x + self.size, y + self.size,
            fill=self.color, outline=""
        )
    
    def move(self):
        coords = self.canvas.coords(self.id)
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        
        if coords[0] <= 0 or coords[2] >= width:
            self.dx = -self.dx
        if coords[1] <= 0 or coords[3] >= height:
            self.dy = -self.dy
            
        self.canvas.move(self.id, self.dx, self.dy)
        return True

# Criar partículas animadas
particles = []
def create_particles():
    global particles
    particles = [Particle(bg_canvas) for _ in range(100)]
    animate_particles()

def animate_particles():
    for particle in particles[:]:
        if not particle.move():
            particles.remove(particle)
    if particles:
        root.after(30, animate_particles)  # Aumentei o tempo entre frames para 30ms

# Atualizar tema
def apply_theme():
    global dark_mode, logo_img
    
    bg_color = DARK_BG if dark_mode else LIGHT_BG
    text_color = DARK_TEXT if dark_mode else LIGHT_TEXT
    button_bg = ACCENT_COLOR
    entry_bg = "#1e1e1e" if dark_mode else "#ffffff"
    
    # Configurar cores básicas
    root.configure(bg=bg_color)
    bg_canvas.configure(bg=bg_color)
    
    # Atualizar estilo dos widgets
    style = ttk.Style()
    style.theme_use("clam")
    
    style.configure(".", 
        background=bg_color,
        foreground=text_color,
        font=(MODERN_FONT, 11)
    )
    
    style.configure("TButton",
        background=button_bg,
        foreground="#ffffff",
        font=(MODERN_FONT, 11, "bold"),
        borderwidth=0,
        relief="flat",
        padding=10
    )
    
    style.map("TButton",
        background=[("active", "#5a7da7")],
        relief=[("active", "flat")]
    )
    
    style.configure("TFrame", background=bg_color)
    style.configure("TLabel", background=bg_color, foreground=text_color)
    style.configure("TEntry", fieldbackground=entry_bg)
    style.configure("TCombobox", fieldbackground=entry_bg)
    
    # Atualizar estilo da barra de progresso
    style.configure("Custom.Horizontal.TProgressbar",
        background=ACCENT_COLOR,
        troughcolor=bg_color,
        bordercolor=bg_color,
        lightcolor=ACCENT_COLOR,
        darkcolor=ACCENT_COLOR,
        thickness=20
    )
    
    # Atualizar cores dos widgets não-ttk
    for widget in root.winfo_children():
        if isinstance(widget, tk.Text):
            widget.configure(
                bg=entry_bg,
                fg=text_color,
                insertbackground=text_color,
                selectbackground=ACCENT_COLOR,
                selectforeground="#ffffff"
            )
        elif isinstance(widget, tk.Canvas):
            widget.configure(bg=bg_color)
    
    # Atualizar botão de tema
    toggle_btn.config(text="☀️" if dark_mode else "🌙")
    
    # Recriar partículas com novas cores
    for particle in particles:
        bg_canvas.delete(particle.id)
    create_particles()

# Alternar tema
def toggle_theme():
    global dark_mode
    dark_mode = not dark_mode
    apply_theme()

# Efeito de hover para botões
def on_enter(e):
    e.widget['background'] = '#5a7da7'

def on_leave(e):
    e.widget['background'] = ACCENT_COLOR

# Animação de clique
def animate_button(button):
    original_bg = button['background']
    button.config(background="#3a5a8a")
    button.after(100, lambda: button.config(background=original_bg))

# Frame principal
main_frame = ttk.Frame(bg_canvas)
main_frame.place(relx=0.5, rely=0.5, anchor="center")

# Botão de tema no canto superior direito
toggle_btn = tk.Button(
    bg_canvas,
    text="🌙",
    command=toggle_theme,
    bg=ACCENT_COLOR,
    fg="white",
    activebackground="#5a7da7",
    activeforeground="white",
    relief="flat",
    font=(MODERN_FONT, 12),
    borderwidth=0,
    padx=10,
    pady=5
)
toggle_btn.place(relx=0.98, rely=0.02, anchor="ne")

# Cabeçalho
header_frame = ttk.Frame(main_frame)
header_frame.pack(pady=(0, 20))

# Logo
logo_img = load_image("C:/GOC/GOC_Launcher/Img/BG/decsis.png", (300, 80))
logo_label = ttk.Label(header_frame, image=logo_img)
logo_label.pack()

# Classe principal da aplicação
class TicketSearchApp:
    def __init__(self, parent):
        self.parent = parent
        self.tickets_jira = []
        self.tickets_pendentes = []
        self.tickets_processar = []
        self.found_emails = []
        self.start_time = 0
        self.estimated_time = 0
        
        # Container principal
        self.container = ttk.Frame(main_frame)
        self.container.pack(fill="both", expand=True)
        
        # Frame de status
        self.status_frame = ttk.Frame(self.container)
        self.status_frame.pack(fill="x", pady=10)
        
        # Status
        self.status_label = ttk.Label(
            self.status_frame,
            text="Status: Pronto",
            font=(MODERN_FONT, 11)
        )
        self.status_label.pack(side="left", padx=5)
        
        # Label de porcentagem
        self.percent_label = ttk.Label(
            self.status_frame,
            text="0%",
            font=(MODERN_FONT, 11)
        )
        self.percent_label.pack(side="right", padx=5)
        
        # Label de tempo estimado
        self.time_label = ttk.Label(
            self.status_frame,
            text="Tempo estimado: --",
            font=(MODERN_FONT, 11)
        )
        self.time_label.pack(side="right", padx=10)
        
        # Barra de progresso
        self.progress_frame = ttk.Frame(self.container)
        self.progress_frame.pack(fill="x", pady=5)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            orient="horizontal",
            mode="determinate",
            style="Custom.Horizontal.TProgressbar",
            length=600
        )
        self.progress_bar.pack(fill="x", expand=True)
        
        # Área de resultados
        self.result_frame = ttk.Frame(self.container)
        self.result_frame.pack(fill="both", expand=True, pady=10)
        
        self.result_text = tk.Text(
            self.result_frame,
            wrap="word",
            font=(MODERN_FONT, 11),
            height=2,
            padx=2,
            pady=2
        )
        self.result_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(
            self.result_frame,
            orient="vertical",
            command=self.result_text.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.result_text.config(yscrollcommand=scrollbar.set)
        
        # Frame de botões
        self.button_frame = ttk.Frame(self.container)
        self.button_frame.pack(fill="x", pady=10)
        
        # Botão Importar
        self.import_btn = tk.Button(
            self.button_frame,
            text="📂 Importar Excel",
            command=self.importar_excel,
            bg=ACCENT_COLOR,
            fg="white",
            activebackground="#5a7da7",
            activeforeground="white",
            relief="flat",
            font=(MODERN_FONT, 11, "bold"),
            borderwidth=0,
            padx=20,
            pady=10
        )
        self.import_btn.pack(side="left", padx=5, fill="x", expand=True)
        self.import_btn.bind("<Enter>", on_enter)
        self.import_btn.bind("<Leave>", on_leave)
        
        # Botão Buscar
        self.search_btn = tk.Button(
            self.button_frame,
            text="🔍 Procurar Emails",
            command=self.buscar_emails,
            bg="#2e7d32",
            fg="white",
            activebackground="#388e3c",
            activeforeground="white",
            relief="flat",
            font=(MODERN_FONT, 11, "bold"),
            borderwidth=0,
            padx=20,
            pady=10
        )
        self.search_btn.pack(side="left", padx=5, fill="x", expand=True)
        self.search_btn.bind("<Enter>", lambda e: self.search_btn.config(bg="#388e3c"))
        self.search_btn.bind("<Leave>", lambda e: self.search_btn.config(bg="#2e7d32"))
        
        # Botão Limpar
        self.clear_btn = tk.Button(
            self.button_frame,
            text="🗑️ Limpar Tudo",
            command=self.reset_app,
            bg="#c62828",
            fg="white",
            activebackground="#d32f2f",
            activeforeground="white",
            relief="flat",
            font=(MODERN_FONT, 11, "bold"),
            borderwidth=0,
            padx=20,
            pady=10
        )
        self.clear_btn.pack(side="left", padx=5, fill="x", expand=True)
        self.clear_btn.bind("<Enter>", lambda e: self.clear_btn.config(bg="#d32f2f"))
        self.clear_btn.bind("<Leave>", lambda e: self.clear_btn.config(bg="#c62828"))
        
        # Frame para exibição dos tickets encontrados
        self.canvas_frame = ttk.Frame(self.container)
        self.canvas_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(self.canvas_frame)
        self.scrollbar = ttk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.ticket_labels_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.ticket_labels_frame, anchor="nw")

        # Configuração importante para o redimensionamento
        self.ticket_labels_frame.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all"),
            width=e.width  # Força o canvas a ter a mesma largura que o frame
        ))

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Rodapé
        footer = ttk.Label(
            self.container,
            text=f"© {datetime.now().year} DECSIS - Desenvolvido por Igor Cunha (GOC)",
            font=(MODERN_FONT, 9)
        )
        footer.pack(side="bottom", pady=(20, 0))
        
        # Aplicar tema inicial
        apply_theme()
    
    def update_progress(self, current, total):
        # Atualizar barra de progresso
        progress = (current / total) * 100
        self.progress_bar["value"] = progress
        self.percent_label.config(text=f"{int(progress)}%")
        
        # Calcular tempo estimado restante
        elapsed_time = time.time() - self.start_time
        if current > 0:
            estimated_total = (elapsed_time / current) * total
            remaining = estimated_total - elapsed_time
            
            # Formatando o tempo restante
            if remaining > 60:
                mins = int(remaining // 60)
                secs = int(remaining % 60)
                self.time_label.config(text=f"Tempo estimado: {mins}m {secs}s")
            else:
                self.time_label.config(text=f"Tempo estimado: {int(remaining)}s")
        
        self.parent.update()
    
    def importar_excel(self):
        animate_button(self.import_btn)
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.tickets_jira = []
                self.tickets_pendentes = []
                self.tickets_processar = []
                
                for _, row in df.iterrows():
                    try:
                        ref = row["Ref"]
                        organization = row["Organization->Name"]
                        title = str(row["Title"])
                        last_update = pd.to_datetime(row["Last update"])
                        
                        # Filtrar tickets Jira/IH
                        if any(keyword in title.upper() for keyword in ["IH", "JIRA"]):
                            self.tickets_jira.append({
                                "ref": ref,
                                "organization": organization,
                                "title": title,
                                "last_update": last_update
                            })
                            continue
                        
                        # Calcular diferença de tempo
                        now = datetime.now()
                        delta = now - last_update
                        
                        if delta.days > 3:
                            self.tickets_processar.append({
                                "ref": ref,
                                "organization": organization,
                                "last_update": last_update,
                                "email": None
                            })
                        else:
                            self.tickets_pendentes.append({
                                "ref": ref,
                                "organization": organization,
                                "last_update": last_update,
                                "delta": delta
                            })
                            
                    except Exception as e:
                        print(f"Erro ao processar linha {_}: {str(e)}")
                
                # Efeito visual de sucesso
                self.import_btn.config(bg="#4caf50")
                self.import_btn.after(300, lambda: self.import_btn.config(bg=ACCENT_COLOR))
                
                messagebox.showinfo(
                    "Sucesso",
                    f"✅ {len(df)} tickets analisados!\n\n"
                    f"• {len(self.tickets_jira)} tickets Jira/IH ignorados\n"
                    f"• {len(self.tickets_processar)} tickets para processar\n"
                    f"• {len(self.tickets_pendentes)} tickets atualizados recentemente"
                )
                
            except Exception as e:
                # Efeito visual de erro
                self.import_btn.config(bg="#f44336")
                self.import_btn.after(300, lambda: self.import_btn.config(bg=ACCENT_COLOR))
                
                messagebox.showerror(
                    "Erro",
                    f"❌ Erro ao importar Excel:\n\n{str(e)}"
                )
    
    def buscar_emails(self):
        if not self.tickets_processar:
            messagebox.showwarning("Aviso", "❌ Nenhum ticket para processar!")
            return
        
        animate_button(self.search_btn)
        self.start_time = time.time()
        self.found_emails = []
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders["GOC"].Folders["Inbox"]
            archive = outlook.Folders["Online Archive - GOC@decservices.eu"].Folders["Clientes"]
            
            tickets_encontrados = []
            total_tickets = len(self.tickets_processar)
            
            self.progress_bar["maximum"] = 100
            self.result_text.config(state="normal")
            self.result_text.delete(1.0, "end")
            
            for idx, ticket in enumerate(self.tickets_processar):
                self.status_label.config(text=f"Buscando: {ticket['ref']} - {ticket['organization']}")
                self.update_progress(idx + 1, total_tickets)
                
                email_encontrado = self.buscar_email(inbox, ticket['ref'])
                
                if not email_encontrado:
                    pasta_cliente = self.encontrar_pasta_cliente(archive, ticket['organization'])
                    if pasta_cliente:
                        email_encontrado = self.buscar_email(pasta_cliente, ticket['ref'])
                
                if email_encontrado:
                    ticket["email"] = email_encontrado
                    tickets_encontrados.append(ticket)
                    self.found_emails.append(email_encontrado)
                    self.result_text.insert("end", f"✅ {ticket['ref']} - {ticket['organization']}\n")
                    self.result_text.see("end")
                    self.parent.update()
            
            self.mostrar_resultados(tickets_encontrados)
            self.mostrar_botoes_resposta()
            
            # Efeito visual de conclusão
            self.search_btn.config(bg="#4caf50")
            self.search_btn.after(300, lambda: self.search_btn.config(bg="#2e7d32"))
            
        except Exception as e:
            self.search_btn.config(bg="#f44336")
            self.search_btn.after(300, lambda: self.search_btn.config(bg="#2e7d32"))
            
            messagebox.showerror(
                "Erro",
                f"❌ Erro na busca de emails:\n\n{str(e)}"
            )
        finally:
            self.status_label.config(text="Status: Concluído")
            self.result_text.config(state="disabled")
            self.time_label.config(text="Tempo estimado: --")
    
    def buscar_email(self, pasta, ticket):
        try:
            items = pasta.Items
            items.Sort("[ReceivedTime]", True)  # Ordenar por data de recebimento, mais recente primeiro
            for item in items:
                if ticket.lower() in item.Subject.lower() or ticket.lower() in item.Body.lower():
                    return item
            return None
        except Exception as e:
            print(f"Erro ao procurar email na pasta {pasta.Name} para o ticket {ticket}: {str(e)}")
            return None
    
    def encontrar_pasta_cliente(self, archive, empresa):
        pasta_correspondente = {
            "BI-SILQUE, PRODUTOS DE COMUNICAÇÃO VISUAL, S.A.": "Bi-silque",
            "Nimco Portugal, Lda.": "NIMCO",
            "Fundação Casa da Música": "Fundação Casa da Musica",
            "FLATLANTIC (ACUINOVA)": "FLATLANTIC",
            "Sonae MC": "Sonae",
            "BRISA": "Brisa",
            "WELLS": "Sonae\\Wells",
            "CIMAC - COMUNIDADE INTERMUNICIPAL DO ALENTEJO CENTRAL": "CIMAC",
            "GRI": "Interno\\GRI",
            "ÁGORA - CULTURA E DESPORTO DO PORTO, S.A.": "Ágora",
            "Sheraton Porto Hotel": "Sheraton Hotel",
            "Pinto & Cruz Gestão, Unipessoal Lda": "Pinto & Cruz",
            "Haco-Etiquetas, S.A": "Haco-Etiquetas",
            "Atobe - Mobility Technology, SA": "Atobe",
            "Câmara Municipal de Santa Maria da Feira": "CM-Feira",
            "Fresenius Kabi Pharma Portugal": "Fresenius",
            "SES IMAGOTAG": "Sonae\\SESIMAGOTAG",
            "WORTEN - EQUIPAMENTOS PARA O LAR S A": "Sonae\\Worten\\Modelos"
        }
        nome_pasta = pasta_correspondente.get(empresa, None)
        if nome_pasta:
            try:
                return archive.Folders[nome_pasta]
            except Exception as e:
                print(f"Erro ao procurar pasta da empresa {empresa}: {str(e)}")
        return None
    
    def mostrar_resultados(self, tickets):
        self.result_text.config(state="normal")
        self.result_text.delete(1.0, "end")
        
        if not tickets:
            self.result_text.insert("end", "Nenhum ticket encontrado ❌\n")
        else:
            self.result_text.insert("end", f"✅ {len(tickets)} tickets encontrados:\n\n")
            
            for ticket in tickets:
                if ticket.get("email"):
                    self.result_text.insert("end", f"🔹 {ticket['ref']} - {ticket['organization']}\n")
                    self.result_text.insert("end", f"   Assunto: {ticket['email'].Subject}\n")
                    self.result_text.insert("end", f"   Data: {ticket['email'].ReceivedTime}\n\n")
        
        self.result_text.config(state="disabled")
    
    def mostrar_botoes_resposta(self):
        """Mostrar os botões de 'Responder 📩' e 'Atualizar iTOP' para cada ticket encontrado"""
        for widget in self.ticket_labels_frame.winfo_children():
            widget.destroy()

        # Seção para tickets Jira
        if self.tickets_jira:
            jira_frame = ttk.LabelFrame(self.ticket_labels_frame, text="🚫 Tickets da Plataforma Jira/IH (Ignorados)")
            jira_frame.pack(fill="x", padx=5, pady=5, ipady=5)
            
            for ticket in self.tickets_jira:
                label_text = f"Ticket: {ticket['ref']} - {ticket['organization']} | Título: {ticket['title']}"
                ttk.Label(jira_frame, text=label_text).pack(anchor="w", padx=5, pady=2)

        # Seção para tickets processados
        if any(ticket.get("email") for ticket in self.tickets_processar):
            processados_frame = ttk.LabelFrame(self.ticket_labels_frame, text="✅ Tickets Processados")
            processados_frame.pack(fill="x", padx=5, pady=5, ipady=5)
            
            for ticket in self.tickets_processar:
                if ticket.get("email"):
                    frame = ttk.Frame(processados_frame)
                    frame.pack(fill="x", padx=5, pady=2)
                    
                    ttk.Label(frame, text=f"Ticket: {ticket['ref']} - {ticket['organization']}").pack(side="left")
                    
                    btn_frame = ttk.Frame(frame)
                    btn_frame.pack(side="right")
                    
                    # Botão Atualizar iTop
                    tk.Button(btn_frame,
                        text="Atualizar iTOP",
                        command=lambda t=ticket: self.atualizar_itop(t['email']),
                        bg="#6a1b9a",
                        fg="white",
                        activebackground="#8e24aa",
                        activeforeground="white",
                        relief="flat",
                        font=(MODERN_FONT, 10, "bold"),
                        borderwidth=0,
                        padx=10,
                        pady=5
                    ).pack(side="right", padx=2)
                    
                    # Botão Responder
                    tk.Button(btn_frame,
                        text="Responder 📩",
                        command=lambda t=ticket: self.reply_all(t['email']),
                        bg=ACCENT_COLOR,
                        fg="white",
                        activebackground="#5a7da7",
                        activeforeground="white",
                        relief="flat",
                        font=(MODERN_FONT, 10, "bold"),
                        borderwidth=0,
                        padx=10,
                        pady=5
                    ).pack(side="right", padx=2)

        # Seção para tickets pendentes
        if self.tickets_pendentes:
            pendentes_frame = ttk.LabelFrame(self.ticket_labels_frame, text="⏳ Tickets Atualizados Recentemente")
            pendentes_frame.pack(fill="x", padx=5, pady=5, ipady=5)
            
            for ticket in self.tickets_pendentes:
                delta = datetime.now() - ticket['last_update']
                remaining = timedelta(days=3) - delta
                
                if remaining.total_seconds() > 0:
                    days = remaining.days
                    hours, rem = divmod(remaining.seconds, 3600)
                    minutes = rem // 60
                    
                    time_str = ""
                    if days > 0:
                        time_str += f"{days} dias "
                    if hours > 0:
                        time_str += f"{hours} horas "
                    if minutes > 0 and days == 0:
                        time_str += f"{minutes} minutos"
                    
                    msg = f"Deverá pedir update daqui a {time_str.strip()}"
                else:
                    msg = "Pronto para atualização"
                
                label_text = f"Ticket: {ticket['ref']} - {ticket['organization']} | {msg}"
                ttk.Label(pendentes_frame, text=label_text).pack(anchor="w", padx=5, pady=2)
    
    def atualizar_itop(self, email):
        """Abrir novo email para atualização no iTOP com o último email enviado HOJE como anexo"""
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # Criar novo email
            new_mail = outlook.CreateItem(0)  # 0 = olMailItem
            
            # Configurar destinatário e assunto
            new_mail.To = "itop.tickets@grupodecsis.eu"
            new_mail.Subject = f"Atualização de ticket - {email.Subject}"
            
            # Obter a pasta Sent Items
            sent_folder = outlook.GetNamespace("MAPI").GetDefaultFolder(5)  # 5 = Sent Items
            
            # Data de hoje para filtro
            hoje = date.today()
            user_address = outlook.Session.CurrentUser.Address
            last_sent_today = None
            
            # Ordenar por data de envio (mais recente primeiro)
            sent_folder.Items.Sort("[SentOn]", True)
            
            for item in sent_folder.Items:
                try:
                    # Verificar se foi enviado hoje e pelo usuário atual
                    if (item.SentOn.date() == hoje and 
                        (item.SenderEmailAddress == user_address or 
                         getattr(item.Sender, 'Address', '') == user_address)):
                        last_sent_today = item
                        break
                except Exception as e:
                    print(f"Erro ao verificar email: {e}")
                    continue
            
            # Adicionar o último email enviado HOJE como anexo
            if last_sent_today:
                try:
                    # Salvar temporariamente o email
                    temp_path = os.path.join(os.environ["TEMP"], "ultimo_email_hoje.msg")
                    last_sent_today.SaveAs(temp_path)
                    new_mail.Attachments.Add(temp_path)
                    
                    # Adicionar mensagem no corpo do email
                    new_mail.Body = f"Foi dada a seguinte resposta ao email do cliente: \n\n" \
                                   f"Data/Hora de Envio: {last_sent_today.SentOn.strftime('%d/%m/%Y %H:%M')}\n" \
                                   f"Título do Email: {last_sent_today.Subject}\n\n" \
                                   f"Envia por: {outlook.Session.CurrentUser.Name}"
                except Exception as e:
                    print(f"Erro ao anexar email: {e}")
                    new_mail.Body = "Não foi possível anexar o último email enviado hoje. Por favor verifique manualmente."
            else:
                new_mail.Body = "Não foram encontrados emails enviados hoje na pasta Sent Items."
            
            # Exibir o email para edição
            new_mail.Display()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível criar o email para iTOP:\n{str(e)}")
    
    def reply_all(self, email):
        """Abrir a janela de 'Reply All' para o email encontrado"""
        try:
            reply = email.ReplyAll()
            reply.Display()  # Abrir a janela de resposta no Outlook para edição
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o email:\n{str(e)}")
    
    def reset_app(self):
        animate_button(self.clear_btn)
        self.tickets_jira = []
        self.tickets_pendentes = []
        self.tickets_processar = []
        self.found_emails = []
        self.result_text.config(state="normal")
        self.result_text.delete(1.0, "end")
        self.result_text.config(state="disabled")
        self.progress_bar["value"] = 0
        self.percent_label.config(text="0%")
        self.time_label.config(text="Tempo estimado: --")
        self.status_label.config(text="Status: Pronto")
        
        # Limpar os botões de resposta
        for widget in self.ticket_labels_frame.winfo_children():
            widget.destroy()
        
        # Efeito visual
        self.clear_btn.config(bg="#4caf50")
        self.clear_btn.after(300, lambda: self.clear_btn.config(bg="#c62828"))

# Inicializar a aplicação
if __name__ == "__main__":
    try:
        # Criar partículas animadas
        create_particles()
        
        # Iniciar aplicação
        app = TicketSearchApp(root)
        
        # Aplicar tema inicial
        apply_theme()
        
        root.mainloop()
    except Exception as e:
        print(f"Erro crítico: {str(e)}")
        input("Pressione Enter para sair...")