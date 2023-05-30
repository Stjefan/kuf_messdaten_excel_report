
import importlib.resources


def greet(recipient):
    """Greet a recipient."""
    template = importlib.resources.read_text("kuf_messdaten_wochenbericht", "greeting.txt")
    print(importlib.resources.files("kuf_messdaten_wochenbericht").joinpath("greeting.txt"))
    print(template)
    return template.format(recipient=recipient)