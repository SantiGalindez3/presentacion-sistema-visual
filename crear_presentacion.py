from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

# --- CONFIGURACIÓN DE COLORES (Basado en tu plantilla rosa/rojo) ---
COLOR_FONDO = RGBColor(253, 235, 235)  # Rosa claro
COLOR_TITULO = RGBColor(139, 0, 0)     # Rojo oscuro
COLOR_TEXTO = RGBColor(80, 0, 0)       # Marrón oscuro
COLOR_BLANCO = RGBColor(255, 255, 255)

# --- CREAR PRESENTACIÓN ---
prs = Presentation()
prs.slide_width = Inches(13.333)  # Formato 16:9
prs.slide_height = Inches(7.5)

def agregar_diapositiva(titulo, puntos, imagen_path=None):
    """Función auxiliar para crear diapositivas con el diseño de la plantilla"""
    slide_layout = prs.slide_layouts[6] # Diseño en blanco
    slide = prs.slides.add_slide(slide_layout)

    # 1. Fondo rosa
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_FONDO

    # 2. Título (Caja roja redondeada simulada con texto)
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(1))
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = titulo
    p.font.size = Pt(36)
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

    # 4. Insertar Imagen (si existe)
    if imagen_path and os.path.exists(imagen_path):
        try:
            slide.shapes.add_picture(imagen_path, Inches(7.2), Inches(1.5), height=Inches(5.5))
        except Exception as e:
            print(f"Error al insertar {imagen_path}: {e}")
    elif imagen_path:
        # Si no existe la imagen, poner un rectángulo gris de placeholder
        shape = slide.shapes.add_shape(1, Inches(7.2), Inches(1.5), Inches(5.5), Inches(5.5))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(200, 200, 200)
        shape.text = f"Falta: {os.path.basename(imagen_path)}"

# --- CONSTRUCCIÓN DE LAS DIAPOSITIVAS ---

# Diapositiva 1: Portada
slide_portada = prs.slides.add_slide(prs.slide_layouts[6])
background = slide_portada.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = COLOR_FONDO
txBox = slide_portada.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(2))
tf = txBox.text_frame
p = tf.add_paragraph()
p.text = "Sistema Visual: Anatomía, Fisiología y Correlación Clínica"
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = COLOR_TITULO
p.alignment = PP_ALIGN.CENTER
p2 = tf.add_paragraph()
p2.text = "De la macroanatomía orbitaria a la optogenética"
p2.font.size = Pt(24)
p2.font.color.rgb = COLOR_TEXTO
p2.alignment = PP_ALIGN.CENTER

# Diapositiva 2: Órbita
agregar_diapositiva(
    "Cavidad Orbitaria y Forámenes Anatómicos",
    [
        "Morfología: Pirámide cuadrangular (base anterior, vértice posterior). Eje divergente 18°.",
        "Techo: Hueso frontal + Ala menor del esfenoides (Fosa lagrimal y fosita troclear).",
        "Piso: Maxilar, cigomático y palatino (Conducto infraorbitario).",
        "Medial: Maxilar, lagrimal, etmoides, esfenoides (Surco y conducto nasolagrimal).",
        "Lateral: Ala mayor del esfenoides + Cigomático.",
        "Vértice: Conducto óptico (N. II y A. oftálmica); Fisura orbitaria superior (N. III, IV, VI, V1 y venas)."
    ],
    "imagenes/imagen4.jpg"
)

# Diapositiva 3: Túnicas
agregar_diapositiva(
    "Túnicas del Ojo: Capa Fibrosa y Vascular (Úvea)",
    [
        "Capa Fibrosa (Externa):",
        "  - Esclerótica: 5/6 posteriores. Tejido conectivo denso. Perforada por lámina cribosa.",
        "  - Córnea: 1/6 anterior. Transparente y avascular. Límite en el limbo esclerocorneal.",
        "Capa Vascular (Úvea):",
        "  - Coroides: Muy vascularizada y pigmentada.",
        "  - Cuerpo Ciliar: Músculo ciliar (acomodación) y procesos ciliares (humor acuoso).",
        "  - Iris: Esfínter (parasimpático/miosis) y Dilatador (simpático/midriasis)."
    ],
    "imagenes/imagen10.jpg"
)

# Diapositiva 4: Medios Transparentes
agregar_diapositiva(
    "Medios Transparentes y Dinámica del Humor Acuoso",
    [
        "Medios de Refracción: Córnea, humor acuoso, cristalino y cuerpo vítreo.",
        "Acomodación:",
        "  - Visión Cercana: Contracción del músculo ciliar -> relajación de la zónula -> cristalino más convexo.",
        "  - Visión Lejana: Relajación del músculo ciliar -> tensión zonalar -> cristalino aplanado.",
        "Humor Acuoso:",
        "  - Secreción: Procesos ciliares (Cámara posterior).",
        "  - Circulación: Pupila -> Cámara anterior.",
        "  - Drenaje: Red trabecular -> Conducto de Schlemm."
    ],
    "imagenes/imagen1.jpg" # Puedes cambiar a imagen12 si prefieres
)

# Diapositiva 5: Anexos
agregar_diapositiva(
    "Aparato de Protección y Aparato Lagrimal",
    [
        "Párpados: Tarsos superior e inferior, ligamentos palpebrales. Glándulas de Meibomio y Moll.",
        "Conjuntiva: Palpebral y ocular (sacos conjuntivales).",
        "Aparato Lagrimal:",
        "  - Glándula lagrimal (fosa superolateral).",
        "  - Vías: Puntos lagrimales -> Conductillos -> Saco lagrimal -> Conducto nasolagrimal -> Meato nasal inferior."
    ],
    "imagenes/imagen7.jpg"
)

# Diapositiva 6: Musculatura
agregar_diapositiva(
    "Músculos Extrínsecos del Globo Ocular",
    [
        "Rectos (4): Superior, Inferior, Medial y Lateral. Origen en Anillo de Zinn.",
        "Oblicuos (2):",
        "  - Superior: Pasa por la tróclea.",
        "  - Inferior: Nace en el piso anteromedial.",
        "Acciones:",
        "  - Recto Lateral: Abducción pura.",
        "  - Recto Medial: Aducción pura.",
        "  - Recto Superior: Elevación, aducción e intorsión.",
        "  - Oblicuo Superior: Intorsión, depresión y abducción."
    ],
    "imagenes/imagen13.jpg"
)

# Diapositiva 7: Inervación
agregar_diapositiva(
    "Inervación Oculomotora (Pares Craneales III, IV y VI)",
    [
        "Nervio Troclear (IV): Inerva al Oblicuo Superior (único con origen dorsal).",
        "Nervio Abducens (VI): Inerva al Recto Lateral.",
        "Nervio Oculomotor (III): Inerva Recto Superior, Inferior, Medial, Oblicuo Inferior y Elevador del párpado.",
        "Componente Vegetativo (III): Fibras parasimpáticas del núcleo de Edinger-Westphal -> Ganglio ciliar -> Esfínter pupilar y músculo ciliar."
    ],
    "imagenes/imagen6.jpg"
)

# Diapositiva 8: Histología
agregar_diapositiva(
    "Arquitectura y Organización Retiniana",
    [
        "Topografía: Disco óptico (punto ciego), Mácula lútea y Fóvea central (máxima agudeza).",
        "Estratificación (10 capas): Epitelio pigmentario -> Conos y bastones -> Membranas limitantes -> Capas nucleares y plexiformes -> Células ganglionares -> Fibras del N. Óptico.",
        "Vía Radial: Fotorreceptor -> Bipolar -> Ganglionar."
    ],
    "imagenes/imagen8.jpg"
)

# Diapositiva 9: Fotorreceptores
agregar_diapositiva(
    "Conos vs. Bastones",
    [
        "Bastones:",
        "  - Visión escotópica (baja luminosidad).",
        "  - Monocromática. Pigmento: Rodopsina.",
        "  - Distribución: Retina periférica.",
        "Conos:",
        "  - Visión fotópica (alta luminosidad).",
        "  - Discriminación cromática (Fotopsinas: azul, verde, rojo).",
        "  - Alta agudeza visual. Distribución: Fóvea central."
    ],
    "imagenes/imagen9.jpg"
)

# Diapositiva 10: Fototransducción
agregar_diapositiva(
    "Cascada Bioquímica de la Luz",
    [
        "1. Fotoisomerización: 11-cis-retinal -> Todo-trans-retinal.",
        "2. Activación: Metarrodopsina activa Transducina (Gt).",
        "3. Efecto: Transducina activa Fosfodiesterasa (PDE) -> Hidrólisis de cGMP.",
        "4. Resultado: Cierre de canales de Na+/Ca2+."
    ]
)

# Diapositiva 11: Respuesta Eléctrica
agregar_diapositiva(
    "Respuesta Eléctrica y Procesamiento Neural",
    [
        "Oscuridad: Despolarización constante (-40 mV) por 'corriente oscura'.",
        "Luz: Cierre de canales -> Hiperpolarización (-70 mV).",
        "Neurotransmisor: Disminución de la liberación de Glutamato.",
        "Vías: Células Bipolares ON (se despolarizan) y OFF (se hiperpolarizan)."
    ]
)

# Diapositiva 12: Vía Óptica
agregar_diapositiva(
    "Trayecto y Decusación de la Vía Óptica",
    [
        "Neuronas: 1.ª Bipolar, 2.ª Ganglionar, 3.ª Cuerpo Geniculado Lateral (CGL).",
        "Nervio Óptico (II): Sale por conducto óptico.",
        "Quiasma Óptico:",
        "  - Fibras Nasales (campo temporal): Decusan.",
        "  - Fibras Temporales (campo nasal): Homolaterales.",
        "Tractos Ópticos: Llegan al CGL del Tálamo."
    ],
    "imagenes/imagen2.jpg"
)

# Diapositiva 13: Corteza
agregar_diapositiva(
    "Radiaciones Ópticas y Corteza Visual Primaria",
    [
        "Radiaciones Ópticas (Gratiolet): Fibras superiores e inferiores (Bucle de Meyer).",
        "Corteza Visual Primaria (V1 / Área 17): Labios del surco calcarino (Lóbulo occipital).",
        "4.ª Neurona: Piramidales y estrelladas.",
        "Áreas Secundarias: 18 y 19 (Integración de movimiento y forma)."
    ],
    "imagenes/imagen3.jpg"
)

# Diapositiva 14: Clínica
agregar_diapositiva(
    "Caso Clínico de Integración: Nictalopia",
    [
        "Presentación: Paciente 35 años, imposibilidad de ver en penumbra. Visión diurna normal.",
        "Correlación: Afección selectiva de Bastones.",
        "Bases Bioquímicas: Carencia de Vitamina A o mutación (Retinitis Pigmentosa) -> Incapacidad de regenerar 11-cis-retinal -> Déficit de Rodopsina."
    ],
    "imagenes/imagen11.jpg"
)

# Diapositiva 15: Cierre
agregar_diapositiva(
    "Vanguardia Científica: Optogenética",
    [
        "Premio Nobel 2026: Karl Deisseroth, Peter Hegemann, Georg Nagel.",
        "Fundamento: Expresar canales fotosensibles (opsinas) en neuronas.",
        "Aplicación: Restaurar visión en degeneraciones retinianas avanzadas.",
        "Frase de Cierre: 'De la anatomía descriptiva en la órbita, al control de neuronas individuales mediante la luz.'"
    ],
    "imagenes/imagen5.jpg"
)

# Guardar archivo
prs.save('Presentacion_Sistema_Visual.pptx')
print("¡Presentación creada con éxito! Revisa el archivo 'Presentacion_Sistema_Visual.pptx'")
