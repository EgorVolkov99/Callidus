import logging
from pathlib import Path
from engine.builders.jinja_env import get_latex_env


class TexBuilder:
    def __init__(self, templates_dir: Path):
        self.env = get_latex_env(templates_dir)

    def build(self, template_name: str, context: dict, output_path: Path) -> Path:
        try:
            template = self.env.get_template(template_name)
            rendered_tex = template.render(**context)

            output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(rendered_tex)

            return output_path
        except Exception as e:
            logging.error(f"Ошибка LaTeX шаблонизации: {e}")
            raise RuntimeError(f"Сбой TexBuilder: {e}")