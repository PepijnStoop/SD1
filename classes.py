from dataclasses import dataclass

@dataclass
class Product:
    magazijn: str
    soort: str
    serienummer: int
    type: str
    id: str