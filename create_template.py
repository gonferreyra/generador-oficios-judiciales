#!/usr/bin/env python3
"""
Recrea la plantilla de constatación con estructura EXACTA del PDF
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from pathlib import Path

OUTPUT_DIR = Path("templates/constatacion")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def create_constatacion_template():
    doc = Document()
    
    # Configurar página A4
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)
    
    # Estilo base con sangría de primera línea
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial Narrow'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0, 0, 0)
    
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = Pt(14)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style.paragraph_format.first_line_indent = Cm(1.25)
    
    style.element.rPr.rFonts.set(qn('w:ascii'), 'Arial Narrow')
    style.element.rPr.rFonts.set(qn('w:hAnsi'), 'Arial Narrow')
    style.element.rPr.rFonts.set(qn('w:cs'), 'Arial Narrow')
    
    def add_para(text, bold=False, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, font_name='Arial Narrow', font_size=Pt(12), indent=True):
        p = doc.add_paragraph()
        p.alignment = alignment
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(14)
        p.paragraph_format.first_line_indent = Cm(1.25) if indent else Cm(0)
        run = p.add_run(text)
        run.font.name = font_name
        run.font.size = font_size
        run.bold = bold
        return p
    
    def add_mixed_para(parts, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True):
        p = doc.add_paragraph()
        p.alignment = alignment
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(14)
        p.paragraph_format.first_line_indent = Cm(1.25) if indent else Cm(0)
        for part in parts:
            text = part[0]
            bold = part[1] if len(part) > 1 else False
            font_name = part[2] if len(part) > 2 else 'Arial Narrow'
            font_size = part[3] if len(part) > 3 else Pt(12)
            run = p.add_run(text)
            run.font.name = font_name
            run.font.size = font_size
            run.bold = bold
        return p
    
    # --- ESTRUCTURA EXACTA (8-9 párrafos) ---
    
    # P1: Fecha (DERECHA, SIN sangría)
    add_para("Córdoba, {{dia}} de {{mes}} de {{anio}}.", bold=True, alignment=WD_ALIGN_PARAGRAPH.RIGHT, indent=False)
    
    # P2: Destinatario (IZQUIERDA, SIN sangría, 3 líneas) - variable {{destinatario_constatacion}}
    add_para("{{destinatario_constatacion}}\nS/D", bold=True, alignment=WD_ALIGN_PARAGRAPH.LEFT, indent=False)
    
    # P3: CUERPO PRINCIPAL - TODO EN UN SOLO PÁRRAFO con sangría
    add_mixed_para([
        ('En los autos caratulados "', False),
        ('{{autos_caratulados}}', True),
        ('", que se tramitan por ante este {{juzgado}} de esta Ciudad, secretaría a cargo de la autorizante se ha resuelto librar a Ud. el presente, a fin de que tan luego de recibido y previo las formalidades de ley, se constituya en el inmueble inscripto a la Matrícula n° {{matricula}} a nombre de {{titular}}; y proceda a ', False),
        ('I) ', True),
        ('Deberá verificar número, calle y barrio en que se emplace y si aquellos primeros no existen, deberá aportar todos los datos que permitan facilitar su ubicación, detallando en lo posible las condiciones del terreno, si está tapiado ó cercado, medidas aproximadas. ', False),
        ('II) ', True),
        ('Si hubiera edificación, características generales de ésta, número de dependencias y en particular el estado de paredes, revoques, pisos, techos, accesorios, etc... ', False),
        ('III) ', True),
        ('Deberá informar sobre la asistencia de servicios en el lugar y en la zona. ', False),
        ('IV) ', True),
        ('En caso de existir edificación, deberá identificar a los ocupantes, consignando el carácter de la ocupación, si fueran inquilinos, monto y términos y si tienen o no contrato; poniendo en conocimiento de los ocupantes que la presente medida se realiza como paso previo a la subasta sobre {{que_se_subasta}}. ', False),
        ('V) ', True),
        ('Proceda a la toma de fotografías del bien constatado (espacios y construcciones existentes), a los fines de dar cumplimiento a la publicidad adecuada ordenada por la reglamentación que rige las subastas electrónicas. Asimismo, deberá dar cumplimiento a lo establecido por la Acordada nº 5 Serie B del S.T. de J.', False),
    ])
    
    # P4: Martillero designado (CON sangría)
    add_para('Se encuentra facultado al diligenciamiento del presente el Martillero {{designado}} y/o quien designe.-')
    
    # P5: Allanamiento (CON sangría)
    add_para('Queda Ud. facultado a allanar domicilio y recurrir al auxilio de la fuerza pública si fuera menester y/o mediare resistencia.-')
    
    # P6: Descripción + Nomenclatura (CON sangría, punto seguido)
    add_mixed_para([
        ('Se hace saber a Ud. que la descripción del inmueble es la siguiente: ', True),
        ('{{descripcion_inmueble}}', False),
        (' Nomenclatura Catastral: ', True),
        ('{{nomenclatura_catastral}}.-', False),
    ])
    
    # P7: Texto de exención (CONDICIONAL - solo si hay {{texto_exencion}})
    # Se agrega como párrafo separado con sangría
    add_para('{{texto_exencion}}')
    
    # P8: Fecho (CON sangría)
    add_para('Fecho, sírvase devolver con lo actuado por la misma vía de su recepción.-')
    
    # P9: DIOS GUARDE A UD (CON sangría, JUSTIFICADO)
    add_para('DIOS GUARDE A UD.-', alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    
    # Guardar
    output_path = OUTPUT_DIR / "tribunal.docx"
    doc.save(output_path)
    print(f"✅ Plantilla creada: {output_path}")
    print(f"   Total párrafos: {len(doc.paragraphs)}")

if __name__ == "__main__":
    create_constatacion_template()