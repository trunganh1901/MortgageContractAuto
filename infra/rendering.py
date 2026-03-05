from docxtpl import DocxTemplate
from config.paths import TEMPLATE_DIR, OUTPUT_DIR
from utils.strings import make_safe_filename

def render(template_cfg, context, seq, name):
    template = TEMPLATE_DIR / template_cfg["docx_file"]

    folder = OUTPUT_DIR / make_safe_filename(name)
    folder.mkdir(exist_ok=True)

    out = folder / f"{seq}_{template_cfg['file_prefix']}_{make_safe_filename(name)}.docx"

    doc = DocxTemplate(template)
    doc.render(context)
    doc.save(out)

    return folder
