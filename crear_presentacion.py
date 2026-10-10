import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import PyPDF2

# --- CONFIGURACIÓN ADAPTADA A TUS ARCHIVOS EN GITHUB ---
PLANTILLA_PATH = "plantilla.pptx"  
# Nombre exacto del PDF tal como lo subiste (con espacios y mayúsculas)
PDF_PATH = "Estructura Exposicion - Anatomia y Fisiologia de la Vision.pdf" 
# Las imágenes están en la raíz, no en una carpeta
CARPETA_IMAGENES = "" 
ARCHIVO_SALIDA = "Presentacion_Final_Sistema_Visual.pptx"

# --- COLORES DE LA PLANTILLA ---
COLOR_FONDO = RGBColor(253, 235, 235)
COLOR_TITULO = RGBColor(139, 0, 0)
COLOR_TEXTO = RGBColor(80, 0, 0)

def extraer_texto_pdf(ruta_pdf):
    """Extrae el texto del PDF (con manejo de errores por si falla)."""
    texto_completo = ""
    try:
        with open(ruta_pdf, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                texto_completo += page.extract_text() + "\n"
    except Exception as e:
        print(f"Advertencia: No se pudo leer el PDF ({e}). Usando texto predefinido.")
        texto_completo = "Contenido de respaldo por si falla la lectura del PDF."
    return texto_completo

def agregar_diapositiva(prs, titulo, puntos, imagen_path=None, es_portada=False):
    """Crea una diapositiva con el diseño de la plantilla."""
    slide_layout = prs.slide_layouts[6] # Diseño en blanco
    slide = prs.slides.add_slide(slide_layout)

    # 1. Fondo rosa
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_FONDO

    # 2. Título
    if es_portada:
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(2))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = titulo
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = COLOR_TITULO
        p.alignment = PP_ALIGN.CENTER
        
        p2 = tf.add_paragraph()
        p2.text = "De la macroanatomía orbitaria a la optogenética"
        p2.font.size = Pt(24)
        p2.font.color.rgb = COLOR_TEXTO
        p2.alignment = PP_ALIGN.CENTER
    else:
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = titulo
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = COLOR_TITULO
        p.alignment = PP_ALIGN.CENTER

        # 3. Puntos de texto
        txBox_body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(6.5), Inches(5.5))
        tf_body = txBox_body.text_frame
        tf_body.word_wrap = True
        
        for punto in puntos:
            p = tf_body.add_paragraph()
            p.text = punto
            p.font.size = Pt(18)
            p.font.color.rgb = COLOR_TEXTO
            p.space_after = Pt(10)

    # 4. Insertar Imagen (Buscando en la raíz con nombres OJO#.png)
    if imagen_path and os.path.exists(imagen_path):
        try:
            slide.shapes.add_picture(imagen_path, Inches(7.2), Inches(1.5), height=Inches(5.5))
        except Exception as e:
            print(f"Error al insertar {imagen_path}: {e}")
    elif imagen_path:
        # Si no encuentra la imagen, pone un cuadro gris
        shape = slide.shapes.add_shape(1, Inches(7.2), Inches(1.5), Inches(5.5), Inches(5.5))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(200, 200, 200)
        shape.text = f"Falta: {os.path.basename(imagen_path)}"

def main():
    if not os.path.exists(PLANTILLA_PATH):
        print(f"Error crítico: No se encuentra {PLANTILLA_PATH}")
        return
    
    prs = Presentation(PLANTILLA_PATH)
    texto_pdf = extraer_texto_pdf(PDF_PATH)
    print("PDF procesado. Generando diapositivas...")

    # --- DIAPOSITIVAS ---
    # Portada
    agregar_diapositiva(prs, "Sistema Visual: Anatomía, Fisiología y Correlación Clínica", [], es_portada=True)

    # Bloque 1
    agregar_diapositiva(prs, "Cavidad Orbitaria y Forámenes Anatómicos", [
        "Morfología: Pirámide cuadrangular (base anterior, vértice posterior). Eje divergente 18°.",
        "Techo: Hueso frontal + Ala menor del esfenoides (Fosa lagrimal y fosita troclear).",
        "Piso: Maxilar, cigomático y palatino (Conducto infraorbitario).",
        "Medial: Maxilar, lagrimal, etmoides, esfenoides (Surco y conducto nasolagrimal).",
        "Lateral: Ala mayor del esfenoides + Cigomático.",
        "Vértice: Conducto óptico (N. II y A. oftálmica); Fisura orbitaria superior (N. III, IV, VI, V1 y venas)."
    ], f"{CARPETA_IMAGENES}OJO4.png") # <--- CAMBIADO A OJO4.png

    agregar_diapositiva(prs, "Túnicas del Ojo: Capa Fibrosa y Vascular (Úvea)", [
        "Capa Fibrosa (Externa):",
        "  - Esclerótica: 5/6 posteriores. Tejido conectivo denso. Perforada por lámina cribosa.",
        "  - Córnea: 1/6 anterior. Transparente y avascular. Límite en el limbo esclerocorneal.",
        "Capa Vascular (Úvea):",
        "  - Coroides: Muy vascularizada y pigmentada.",
        "  - Cuerpo Ciliar: Músculo ciliar (acomodación) y procesos ciliares (humor acuoso).",
        "  - Iris: Esfínter (parasimpático/miosis) y Dilatador (simpático/midriasis)."
    ], f"{CARPETA_IMAGENES}OJO10.png") # <--- CAMBIADO A OJO10.png

    # Bloque 2
    agregar_diapositiva(prs, "Medios Transparentes y Dinámica del Humor Acuoso", [
        "Medios de Refracción: Córnea, humor acuoso, cristalino y cuerpo vítreo.",
        "Acomodación:",
        "  - Visión Cercana: Contracción del músculo ciliar -> relajación de la zónula -> cristalino más convexo.",
        "  - Visión Lejana: Relajación del músculo ciliar -> tensión zonalar -> cristalino aplanado.",
        "Humor Acuoso:",
        "  - Secreción: Procesos ciliares (Cámara posterior).",
        "  - Circulación: Pupila -> Cámara anterior.",
        "  - Drenaje: Red trabecular -> Conducto de Schlemm."
    ], f"{CARPETA_IMAGENES}OJO1.png") # <--- CAMBIADO A OJO1.png

    agregar_diapositiva(prs, "Aparato de Protección y Aparato Lagrimal", [
        "Párpados: Tarsos superior e inferior, ligamentos palpebrales. Glándulas de Meibomio y Moll.",
        "Conjuntiva: Palpebral y ocular (sacos conjuntivales).",
        "Aparato Lagrimal:",
        "  - Glándula lagrimal (fosa superolateral).",
        "  - Vías: Puntos lagrimales -> Conductillos -> Saco lagrimal -> Conducto nasolagrimal -> Meato nasal inferior."
    ], f"{CARPETA_IMAGENES}OJO7.png") # <--- CAMBIADO A OJO7.png

    # Bloque 3
    agregar_diapositiva(prs, "Músculos Extrínsecos del Globo Ocular", [
        "Rectos (4): Superior, Inferior, Medial y Lateral. Origen en Anillo de Zinn.",
        "Oblicuos (2): Superior (tróclea) e Inferior (piso anteromedial).",
        "Acciones:",
        "  - Recto Lateral: Abducción pura.",
        "  - Recto Medial: Aducción pura.",
        "  - Recto Superior: Elevación, aducción e intorsión.",
        "  - Oblicuo Superior: Intorsión, depresión y abducción."
    ], f"{CARPETA_IMAGENES}OJO13.png") # <--- CAMBIADO A OJO13.png

    agregar_diapositiva(prs, "Inervación Oculomotora (Pares Craneales III, IV y VI)", [
        "Nervio Troclear (IV): Inerva al Oblicuo Superior (único con origen dorsal).",
        "Nervio Abducens (VI): Inerva al Recto Lateral.",
        "Nervio Oculomotor (III): Inerva Recto Superior, Inferior, Medial, Oblicuo Inferior y Elevador del párpado.",
        "Componente Vegetativo (III): Fibras parasimpáticas del núcleo de Edinger-Westphal -> Ganglio ciliar -> Esfínter pupilar y músculo ciliar."
    ], f"{CARPETA_IMAGENES}OJO6.png") # <--- CAMBIADO A OJO6.png

    # Bloque 4
    agregar_diapositiva(prs, "Arquitectura y Organización Retiniana", [
        "Topografía: Disco óptico (punto ciego), Mácula lútea y Fóvea central (máxima agudeza).",
        "Estratificación (10 capas): Epitelio pigmentario -> Conos y bastones -> Membranas limitantes -> Capas nucleares y plexiformes -> Células ganglionares -> Fibras del N. Óptico.",
        "Vía Radial: Fotorreceptor -> Bipolar -> Ganglionar."
    ], f"{CARPETA_IMAGENES}OJO8.png") # <--- CAMBIADO A OJO8.png

    agregar_diapositiva(prs, "Conos vs. Bastones", [
        "Bastones: Visión escotópica (baja luminosidad). Monocromática. Rodopsina. Retina periférica.",
        "Conos: Visión fotópica (alta luminosidad). Discriminación cromática (Fotopsinas: azul, verde, rojo). Alta agudeza visual. Fóvea central."
    ], f"{CARPETA_IMAGENES}OJO9.png") # <--- CAMBIADO A OJO9.png

    # Bloque 5
    agregar_diapositiva(prs, "Cascada Bioquímica de la Luz (Fototransducción)", [
        "1. Fotoisomerización: 11-cis-retinal -> Todo-trans-retinal.",
        "2. Activación: Metarrodopsina activa Transducina (Gt).",
        "3. Efecto: Transducina activa Fosfodiesterasa (PDE) -> Hidrólisis de cGMP.",
        "4. Resultado: Cierre de canales de Na+/Ca2+."
    ]) # Sin imagen específica

    agregar_diapositiva(prs, "Respuesta Eléctrica y Procesamiento Neural", [
        "Oscuridad: Despolarización constante (-40 mV) por 'corriente oscura'.",
        "Luz: Cierre de canales -> Hiperpolarización (-70 mV).",
        "Neurotransmisor: Disminución de la liberación de Glutamato.",
        "Vías: Células Bipolares ON (se despolarizan) y OFF (se hiperpolarizan)."
    ]) # Sin imagen específica

    # Bloque 6
    agregar_diapositiva(prs, "Trayecto y Decusación de la Vía Óptica", [
        "Neuronas: 1.ª Bipolar, 2.ª Ganglionar, 3.ª Cuerpo Geniculado Lateral (CGL).",
        "Nervio Óptico (II): Sale por conducto óptico.",
        "Quiasma Óptico: Fibras Nasales (campo temporal) decusan; Fibras Temporales (campo nasal) homolaterales.",
        "Tractos Ópticos: Llegan al CGL del Tálamo."
    ], f"{CARPETA_IMAGENES}OJO2.png") # <--- CAMBIADO A OJO2.png

    agregar_diapositiva(prs, "Radiaciones Ópticas y Corteza Visual Primaria", [
        "Radiaciones Ópticas (Gratiolet): Fibras superiores e inferiores (Bucle de Meyer).",
        "Corteza Visual Primaria (V1 / Área 17): Labios del surco calcarino (Lóbulo occipital).",
        "4.ª Neurona: Piramidales y estrelladas.",
        "Áreas Secundarias: 18 y 19 (Integración de movimiento y forma)."
    ], f"{CARPETA_IMAGENES}OJO3.png") # <--- CAMBIADO A OJO3.png

    # Bloque 7
    agregar_diapositiva(prs, "Caso Clínico de Integración: Nictalopia", [
        "Presentación: Paciente 35 años, imposibilidad de ver en penumbra. Visión diurna normal.",
        "Correlación: Afección selectiva de Bastones.",
        "Bases Bioquímicas: Carencia de Vitamina A o mutación (Retinitis Pigmentosa) -> Incapacidad de regenerar 11-cis-retinal -> Déficit de Rodopsina."
    ], f"{CARPETA_IMAGENES}OJO11.png") # <--- CAMBIADO A OJO11.png

    agregar_diapositiva(prs, "Vanguardia Científica: Optogenética", [
        "Premio Nobel 2026: Karl Deisseroth, Peter Hegemann, Georg Nagel.",
        "Fundamento: Expresar canales fotosensibles (opsinas) en neuronas.",
        "Aplicación: Restaurar visión en degeneraciones retinianas avanzadas.",
        "Frase de Cierre: 'De la anatomía descriptiva en la órbita, al control de neuronas individuales mediante la luz.'"
    ], f"{CARPETA_IMAGENES}OJO5.png") # <--- CAMBIADO A OJO5.png

    # Guardar
    prs.save(ARCHIVO_SALIDA)
    print(f"¡Éxito! Presentación guardada como: {ARCHIVO_SALIDA}")

if __name__ == "__main__":
    main()
