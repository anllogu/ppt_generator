import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os

class SuccessStoryGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Success Story CRM")
        self.root.geometry("900x700")
        
        # Variables para almacenar datos
        self.template_path = tk.StringVar()
        self.client_name = tk.StringVar()
        self.sector = tk.StringVar()
        self.tech = tk.StringVar()
        self.country = tk.StringVar()
        
        # Crear interfaz
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Sección para seleccionar la plantilla
        template_frame = ttk.LabelFrame(main_frame, text="Plantilla PowerPoint", padding="10")
        template_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(template_frame, text="Selecciona la plantilla PPTX:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(template_frame, textvariable=self.template_path, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(template_frame, text="Examinar", command=self.browse_template).grid(row=0, column=2)
        
        # Sección para datos del encabezado
        header_frame = ttk.LabelFrame(main_frame, text="Datos del Encabezado", padding="10")
        header_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(header_frame, text="Cliente:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(header_frame, textvariable=self.client_name, width=30).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(header_frame, text="Sector:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(header_frame, textvariable=self.sector, width=30).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(header_frame, text="Tecnología:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(header_frame, textvariable=self.tech, width=30).grid(row=2, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(header_frame, text="País:").grid(row=3, column=0, sticky=tk.W, pady=2)
        ttk.Entry(header_frame, textvariable=self.country, width=30).grid(row=3, column=1, sticky=tk.W, pady=2)
        
        # Sección para título y subtítulo
        title_frame = ttk.LabelFrame(main_frame, text="Título y Subtítulo", padding="10")
        title_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(title_frame, text="Título:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.title_entry = ttk.Entry(title_frame, width=80)
        self.title_entry.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.title_entry.insert(0, "Success Story @ CRM")
        
        ttk.Label(title_frame, text="Subtítulo:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.subtitle_entry = ttk.Entry(title_frame, width=80)
        self.subtitle_entry.grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # Sección para contenido
        content_frame = ttk.LabelFrame(main_frame, text="Contenido", padding="10")
        content_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        content_notebook = ttk.Notebook(content_frame)
        content_notebook.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña Client Challenge
        challenge_frame = ttk.Frame(content_notebook, padding="10")
        content_notebook.add(challenge_frame, text="Client Challenge")
        
        ttk.Label(challenge_frame, text="Añade cada punto en una línea separada:").pack(anchor=tk.W)
        self.challenge_text = scrolledtext.ScrolledText(challenge_frame, height=10)
        self.challenge_text.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña Business Value
        value_frame = ttk.Frame(content_notebook, padding="10")
        content_notebook.add(value_frame, text="Business Value")
        
        ttk.Label(value_frame, text="Añade cada punto en una línea separada:").pack(anchor=tk.W)
        self.value_text = scrolledtext.ScrolledText(value_frame, height=10)
        self.value_text.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña Solution
        solution_frame = ttk.Frame(content_notebook, padding="10")
        content_notebook.add(solution_frame, text="Solution")
        
        ttk.Label(solution_frame, text="Añade cada punto en una línea separada:").pack(anchor=tk.W)
        self.solution_text = scrolledtext.ScrolledText(solution_frame, height=10)
        self.solution_text.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña Technical Information
        tech_frame = ttk.Frame(content_notebook, padding="10")
        content_notebook.add(tech_frame, text="Technical Information")
        
        ttk.Label(tech_frame, text="Añade cada punto en una línea separada:").pack(anchor=tk.W)
        self.tech_info_text = scrolledtext.ScrolledText(tech_frame, height=10)
        self.tech_info_text.pack(fill=tk.BOTH, expand=True)
        
        # Botones de acción
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="Generar PowerPoint", command=self.generate_ppt).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Limpiar", command=self.clear_fields).pack(side=tk.RIGHT, padx=5)
    
    def browse_template(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PowerPoint files", "*.pptx")],
            title="Seleccionar plantilla PowerPoint"
        )
        if file_path:
            self.template_path.set(file_path)
    
    def clear_fields(self):
        self.client_name.set("")
        self.sector.set("")
        self.tech.set("")
        self.country.set("")
        self.subtitle_entry.delete(0, tk.END)
        self.challenge_text.delete(1.0, tk.END)
        self.value_text.delete(1.0, tk.END)
        self.solution_text.delete(1.0, tk.END)
        self.tech_info_text.delete(1.0, tk.END)
    
    def format_bullet_points(self, text_content):
        """Formatea el texto para crear puntos con viñetas"""
        lines = text_content.strip().split('\n')
        bullet_points = []
        
        for line in lines:
            if line.strip():
                # Si la línea no empieza con un punto de viñeta, añadirlo
                if not line.strip().startswith('•'):
                    bullet_points.append(f"• {line.strip()}")
                else:
                    bullet_points.append(line.strip())
        
        return bullet_points
    
    def find_shape_by_name(self, slide, shape_name):
        """Busca una forma por su nombre en la diapositiva"""
        for shape in slide.shapes:
            if shape.name == shape_name:
                return shape
        return None
    
    def find_textbox_by_text_contains(self, slide, text_contains):
        """Busca un cuadro de texto que contenga cierto texto"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                if text_contains.lower() in shape.text.lower():
                    return shape
        return None
    
    def insert_content_to_placeholder(self, shape, bullet_points):
        """Inserta contenido con formato de viñetas en un placeholder"""
        if not shape or not shape.has_text_frame:
            return False
        
        text_frame = shape.text_frame
        text_frame.clear()  # Limpia el contenido existente
        
        for i, point in enumerate(bullet_points):
            p = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
            p.text = point
            p.level = 0
        
        return True
    
    def generate_ppt(self):
        template_path = self.template_path.get()
        
        if not template_path:
            messagebox.showerror("Error", "Por favor selecciona una plantilla PowerPoint")
            return
        
        if not os.path.exists(template_path):
            messagebox.showerror("Error", "La plantilla seleccionada no existe")
            return
        
        try:
            # Cargar la presentación existente
            prs = Presentation(template_path)
            
            # Formatear los textos de entrada
            challenge_points = self.format_bullet_points(self.challenge_text.get(1.0, tk.END))
            value_points = self.format_bullet_points(self.value_text.get(1.0, tk.END))
            solution_points = self.format_bullet_points(self.solution_text.get(1.0, tk.END))
            tech_info_points = self.format_bullet_points(self.tech_info_text.get(1.0, tk.END))
            
            # Actualizar slide con los datos ingresados
            if prs.slides:
                slide = prs.slides[0]  # Usar la primera diapositiva
                
                # Actualizar título y subtítulo si existen
                title_shape = self.find_textbox_by_text_contains(slide, "Success Story")
                if title_shape and title_shape.has_text_frame:
                    title_shape.text_frame.paragraphs[0].text = self.title_entry.get()
                
                subtitle_shape = self.find_textbox_by_text_contains(slide, "Optimización")
                if subtitle_shape and subtitle_shape.has_text_frame:
                    subtitle_shape.text_frame.paragraphs[0].text = self.subtitle_entry.get()
                
                # Actualizar datos del cliente
                client_shape = self.find_textbox_by_text_contains(slide, "VolksWagen")
                if client_shape and client_shape.has_text_frame:
                    client_shape.text_frame.paragraphs[0].text = self.client_name.get()
                
                sector_shape = self.find_textbox_by_text_contains(slide, "Automan")
                if sector_shape and sector_shape.has_text_frame:
                    sector_shape.text_frame.paragraphs[0].text = self.sector.get()
                
                tech_shape = self.find_textbox_by_text_contains(slide, "Salesforce")
                if tech_shape and tech_shape.has_text_frame:
                    tech_shape.text_frame.paragraphs[0].text = self.tech.get()
                
                country_shape = self.find_textbox_by_text_contains(slide, "Spain")
                if country_shape and country_shape.has_text_frame:
                    country_shape.text_frame.paragraphs[0].text = self.country.get()
                
                # Actualizar secciones principales
                challenge_shape = self.find_textbox_by_text_contains(slide, "CLIENT CHALLENGE")
                if challenge_shape:
                    # Buscar el contenedor de los puntos (generalmente está después del título)
                    for shape in slide.shapes:
                        if shape.has_text_frame and "Lentitud en la generación" in shape.text:
                            self.insert_content_to_placeholder(shape, challenge_points)
                            break
                
                value_shape = self.find_textbox_by_text_contains(slide, "BUSINESS VALUE")
                if value_shape:
                    for shape in slide.shapes:
                        if shape.has_text_frame and "Reducción del tiempo" in shape.text:
                            self.insert_content_to_placeholder(shape, value_points)
                            break
                
                solution_shape = self.find_textbox_by_text_contains(slide, "SOLUTION")
                if solution_shape:
                    for shape in slide.shapes:
                        if shape.has_text_frame and "Desarrollo de una solución" in shape.text:
                            self.insert_content_to_placeholder(shape, solution_points)
                            break
                
                tech_info_shape = self.find_textbox_by_text_contains(slide, "TECHNICAL INFORMATION")
                if tech_info_shape:
                    for shape in slide.shapes:
                        if shape.has_text_frame and "Agentes comerciales" in shape.text:
                            self.insert_content_to_placeholder(shape, tech_info_points)
                            break
            
            # Mostrar diálogo para guardar archivo
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pptx",
                filetypes=[("PowerPoint files", "*.pptx")],
                title="Guardar presentación como"
            )
            
            if file_path:
                prs.save(file_path)
                messagebox.showinfo("Éxito", f"Presentación guardada en:\n{file_path}")
                
                # Preguntar si desea abrir el archivo
                if messagebox.askyesno("Abrir archivo", "¿Deseas abrir la presentación?"):
                    os.startfile(file_path)
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar la presentación: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SuccessStoryGenerator(root)
    root.mainloop()