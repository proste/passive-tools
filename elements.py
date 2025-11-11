import re

import numpy as np


def parse_spec_str(spec: str) -> dict:
    matches = re.findall(r'([a-zA-Z\d,]+)=(\d+)', spec)
    return {k_: float(v) for k, v in matches for k_ in k.split(",")}


class Pricelist:
    sheet_metal_m2 = 350
    flange = 220
    pipe_metal_m2 = 585
    pipe_fitting_metal_m2 = 814
    pipe_fitting_piece = 105
    
    def __init__(self):
        pass

    def __getitem__(self, key):
        raise KeyError()


class BaseElement:
    """
    Abstract Base Class for a single element from the CAD export.
    Price aside, it gets fully defined upon initialization.
    """
    
    EXTRA_ATTRS = ()
    UNIT = None
    REGISTERED_ELEMENTS: list[type["BaseElement"]] = []

    def __init_subclass__(cls, **kwargs):
        """Automatically register all subclasses with the factory."""
        super().__init_subclass__(**kwargs)
        if cls not in cls.REGISTERED_ELEMENTS:
            cls.REGISTERED_ELEMENTS.append(cls)

    def __init__(self, row: dict, pricelist: Pricelist):
        self.issues: list[str | tuple[str, Exception]] = []

        # 1. Parse common properties
        row = dict(row)
        self.system = row.pop('system')
        if not self.system:
            self.issues.append("chybí systém")
            self.system = "SYSTÉM"
        self.position = row.pop("position")
        if not self.position:
            self.issues.append("chybí pozice")
            self.position = 'POZICE'
        self.pn = row.pop('pn')
        self.name = row.pop('name')
        self.spec = row.pop('spec')
        self.quantity = row.pop('quantity')
        if not (self.quantity > 0):
            self.issues.append("chybějící/nulové množství")
        self.unit = row.pop('unit')
        self.insulation_mm = row.pop("insulation_mm")

        if self.UNIT:
            if self.unit != self.UNIT:
                self.issues.append(f"jednotka má být [{self.UNIT}]")
        elif self.unit not in {"ks", "m", "m2"}:
            self.issues.append("špatná jednotka")
            

        # 2. Set and parse extra properties
        row = {k: v for k, v in row.items() if not np.isnan(v)}
        if missing_attrs := (set(self.EXTRA_ATTRS) - set(row)):
            self.issues.append(f"chybi {', '.join(missing_attrs)}")
            for attr_name in missing_attrs:
                row[attr_name] = 0
        for k, v in row.items():
            if not hasattr(self, k):
                setattr(self, k, v)
        try:
            self._parse_spec()
        except Exception as ex:
            self.issues.append(("selhalo vyčtení specifikace prvku", ex))
        self.extra_attrs = row
        
        # 3. Calculate insulation
        if self.insulation_mm > 0:
            try:
                self.insulation_area_m2 = self._calculate_insulation_mm2() / 1_000_000
            except Exception as ex:
                self.issues.append(("selhal výpočet izolace", ex))
                self.insulation_area_m2 = getattr(self, "surface_m2", 0)
        else:
            self.insulation_area_m2 = np.nan

        # 3. Calculate unit_price
        try:
            self.price = self._calculate_price(pricelist)
        except Exception as ex:
            self.issues.append(("selhal výpočet ceny", ex))
            self.price = np.nan

    def __repr__(self):
        if self.insulation_mm:
            insulation = f"{round(self.insulation_area_m2, 2)}m2 @ {round(self.insulation_mm)}mm; "
        else:
            insulation = ""
        return f"{self.__class__.__name__}({self.name}; {self.spec}; {self.quantity}{self.unit}; {insulation}{self.price}CZK)"

    def _parse_spec(self):
        """Parse extra properties."""
        pass

    def _calculate_insulation_mm2(self) -> float:
        """Hook for subclasses. Base implementation returns 0."""
        raise NotImplementedError()

    def _calculate_price(self, pricelist: Pricelist) -> float:
        """
        Base price calculation.
        Subclasses can override this for more complex calculations.
        """
        try:
            unit_price, unit = pricelist[self]
        except:
            return np.nan

        if unit != self.unit:
            self.issues.append("neodpovídá jednotka v ceníku")
            return np.nan

        return unit_price * self.quantity

    def to_dict(self) -> dict:
        """Export element data to a dictionary for final DataFrame summarization."""
        return {
            'system': self.system,
            'position': self.position,
            'pn': self.pn,
            'name': self.name,
            'spec': self.spec,
            'quantity': self.quantity,
            'unit': self.unit,
            'insulation_mm': self.insulation_mm,
            'insulation_area_m2': self.insulation_area_m2,
            'price': self.price,
            'issues': "; ".join((i if isinstance(i, str) else i[0]) for i in self.issues),
        }

    @classmethod
    def can_parse(cls, row: dict) -> bool:
        """
        Factory method: Does this class know how to handle this row?
        The BaseElement handles *nothing*, acting as a fallback.
        """
        return False


class RoundTube(BaseElement):
    """Round tube.

    Has diameter, insulation and surface; measured in meters; cut ad-hoc
    """

    EXTRA_ATTRS = ('diameter_mm', 'surface_m2')
    UNIT = 'm'
    
    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return row['name'] == 'Roura'

    @property
    def length_mm(self) -> float:
        # should be row["duct_count"] * row["length_mm"]
        return 1000 * self.quantity

    def _calculate_insulation_mm2(self) -> float:
        """Hollow cylinder ~ pi * d * l"""
        return np.pi * (self.diameter_mm + 2 * self.insulation_mm) * self.length_mm

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.pipe_metal_m2


class DampedRoundTube(BaseElement):
    """Round tube, wrapped in sound damping material.

    Has diameter, length, insulation and acoustic insulation; measured in pieces; ready made
    """

    EXTRA_ATTRS = ('width_mm', 'height_mm', 'length_mm', 'surface_m2')
    UNIT = 'ks'
    
    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return row['name'] == 'Tlumič hluku, kulatý'

    def _parse_spec(self):
        if self.width_mm != self.height_mm:
            self.issues.append("nejednoznačný průměr")
        self.diameter_mm = self.width_mm
        del self.width_mm
        del self.height_mm

        match = re.search(r'\d+/\d+/(\d+)', self.spec)
        try:
            self.acoustic_mm = int(match.group(1))
        except Exception as ex:
            self.issues.append(("selhalo vyčtení tloušťky akustické izolace", ex))
        
    def _calculate_insulation_mm2(self) -> float:
        """Hollow cylinder ~ q * pi * d * l"""
        circumference = np.pi * (self.diameter_mm + 2 * (self.insulation_mm + self.acoustic_mm))
        return self.quantity * circumference * self.length_mm


class RoundTubeJoint(BaseElement):
    """Round tube inserted into two tubes.

    Has diameter and surface; measured in meters; cut ad-hoc
    """

    EXTRA_ATTRS = ('diameter_mm', 'surface_m2')
    UNIT = 'ks'
    
    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return row['name'] in 'Vsuvka do potrubí'

    def _calculate_insulation_mm2(self) -> float:
        """Completely encapsulated by other elements"""
        return 0

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.pipe_metal_m2


class FlatTube(BaseElement):
    """Flat tube with flanges.

    Has width, height, length, insulation and surface; measured in pieces; custom-made
    """

    EXTRA_ATTRS = ('width_mm', 'height_mm', 'length_mm', 'duct_count', 'surface_m2')
    UNIT = 'm'
    
    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return row['name'] == 'Potrubí'

    def _parse_spec(self):
        self.quantity = self.duct_count
        del self.duct_count
        self.unit = 'ks'
        self.spec = f"{self.spec} x {int(self.length_mm)}"

    def _calculate_insulation_mm2(self) -> float:
        """Hollow box ~ quantity * 2 * (w + h) * l"""
        circumference = 2 * (self.width_mm + self.height_mm + 4 * self.insulation_mm)
        return self.quantity * circumference * self.length_mm

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.sheet_metal_m2 + 2 * self.quantity * pricelist.flange


class FloorFlatTube(FlatTube):
    """Flat tube.

    Has width, height, length, insulation and surface; measured in metres; cut ad-hoc
    """

    EXTRA_ATTRS = ('width_mm', 'height_mm', 'surface_m2')
    UNIT = 'm'
    
    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return (
            super().can_parse(row)
            and ((row['width_mm'], row['height_mm']) in [(160, 40), (200, 50)])
        )

    @property
    def length_mm(self) -> float:
        # should be row["duct_count"] * row["length_mm"]
        return 1000 * self.quantity

    def _parse_spec(self):
        self.name = "Podlahový kanál"

    def _calculate_insulation_mm2(self) -> float:
        """Hollow box ~ 2 * (w + h) * l"""
        return 2 * (self.width_mm + self.height_mm + 4 * self.insulation_mm) * self.length_mm

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.sheet_metal_m2


class DampedFlatTube(FlatTube):
    """Flat tube, with sound damping elements inside.

    Has width, height, length, insulation and surface; measured in pieces; ready made
    """

    EXTRA_ATTRS = ('width_mm', 'height_mm', 'length_mm', 'surface_m2')
    UNIT = 'ks'

    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return row['name'] == 'Tlumič hluku, buňkový'

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return BaseElement._calculate_price(pricelist)


class RoundElbow(BaseElement):
    """Curved round tube.

    Has diameter, radius, angle, insulation and surface; measured in pieces; ready made
    """

    EXTRA_ATTRS = ('diameter_mm', 'surface_m2')
    UNIT = 'ks'
    
    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return (row['name'] == 'Koleno') and ("D=" in row['spec'])

    def _parse_spec(self):
        spec = parse_spec_str(self.spec)
        self.radius_mm = spec['R']
        self.angle_deg = spec['a']

    def _calculate_insulation_mm2(self) -> float:
        """Hollow cylindrical arc ~ q * (360 / a) * (2 * pi * r) * (pi * d)"""
        arc_len = (self.angle_deg / 360) * (2 * np.pi * (self.radius_mm + self.diameter_mm / 2 + self.insulation_mm))
        circumference = np.pi * (self.diameter_mm + 2 * self.insulation_mm)
        return self.quantity * arc_len * circumference

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.pipe_fitting_metal_m2 + pricelist.pipe_fitting_piece


class FlatElbow(BaseElement):
    """Curved flat tube with flanges.

    Has width, height, radius, angle, insulation and surface; measured in pieces; ready made
    """

    EXTRA_ATTRS = ('width_mm', 'height_mm', 'surface_m2')
    UNIT = 'ks'

    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return (row['name'] == 'Koleno') and ("A=" in row['spec'])

    def _parse_spec(self):
        spec = parse_spec_str(self.spec)
        self.radius_mm = spec['R']
        self.angle_deg = spec['a']

    def _calculate_insulation_mm2(self) -> float:
        """Hollow flat arc ~ q * (360 / a) * (2 * pi * r) * 2 * (a + b)"""
        arc_len = (self.angle_deg / 360) * (2 * np.pi * (self.radius_mm + self.width_mm + self.insulation_mm))
        circumference = 2 * (self.width_mm + self.height_mm + 4 * self.insulation_mm)
        return self.quantity * arc_len * circumference

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.sheet_metal_m2 + 2 * pricelist.flange


class FlatReduction(BaseElement):
    """Flat/Flat tube reduction with flanges.

    Has two sets of width and height, length and surface; measured in pieces; ready made
    """

    EXTRA_ATTRS = ('length_mm', 'surface_m2')
    UNIT = 'ks'

    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return (row['name'] == 'Redukce') and ("A=" in row['spec'])

    def _parse_spec(self):
        spec = parse_spec_str(self.spec)
        self.width_mm = max(spec['A'], spec.get('A2', 0))
        self.height_mm = max(spec['B'], spec.get('B2', 0))

    def _calculate_insulation_mm2(self) -> float:
        """Hollow flat trapezoid, approx as convex hull - box ~ q * 2 * (a + b) * l"""
        circumference = 2 * (self.width_mm + self.height_mm + 4 * self.insulation_mm)
        return self.quantity * circumference * self.length_mm

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.sheet_metal_m2 + 2 * pricelist.flange


class RoundReduction(BaseElement):
    """Round/Round tube reduction

    Has two diameters, length and surface; measured in pieces; ready made
    """

    EXTRA_ATTRS = ('length_mm', 'surface_m2',)
    UNIT = 'ks'

    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return (row['name'] == 'Redukce') and ("D=" in row['spec'])

    def _parse_spec(self):
        spec = parse_spec_str(self.spec)
        self.diameter_mm = max(spec['D'], spec['D2'])

    def _calculate_insulation_mm2(self) -> float:
        """Hollow frustum, approx as convex hull - cylinder ~ q * (pi * d) * l"""
        circumference = np.pi * (self.diameter_mm + 2 * self.insulation_mm)
        return self.quantity * circumference * self.length_mm

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.pipe_metal_m2


class FlatRoundReduction(BaseElement):
    """Flat/Round tube reduction

    Has width, height, diameter, length and surface; measured in pieces; ready made
    """

    EXTRA_ATTRS = ('diameter_mm', 'width_mm', 'height_mm', 'length_mm', 'surface_m2')
    UNIT = 'ks'

    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return row['name'] == 'Redukce obdélník-roura'

    def _calculate_insulation_mm2(self) -> float:
        """Hollow frustum, approx as convex hull - cylinder ~ q * (pi * d) * l"""
        circumference = max(
            np.pi * (self.diameter_mm + 2 * self.insulation_mm),
            2 * (self.width_mm + self.height_mm + 4 * self.insulation_mm),
        )
        return self.quantity * circumference * self.length_mm

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * max(pricelist.pipe_fitting_metal_m2, pricelist.sheet_metal_m2) + pricelist.flange + pricelist.pipe_fitting_piece


class RoundTee(BaseElement):
    """Round tee reduction

    Has two diameters, height, diameter, length and surface; measured in pieces; ready made
    """

    EXTRA_ATTRS = ('diameter_mm', 'length_mm', 'surface_m2')
    UNIT = 'ks'

    @classmethod
    def can_parse(cls, row: dict) -> bool:
        return row['name'] == 'T-kus'

    def _parse_spec(self):
        spec = parse_spec_str(self.spec)
        self.diameter3_mm = spec['D3']
        self.length3_mm = spec['L3']
        self.angle_deg = spec['a']

    def _calculate_insulation_mm2(self) -> float:
        """T-shape, approx as two cylinders - cylinder ~ q * (pi * d) * l"""
        main = np.pi * (self.diameter_mm + 2 * self.insulation_mm) * self.length_mm
        aux = np.pi * (self.diameter3_mm + 2 * self.insulation_mm) * (self.length3_mm - self.diameter_mm / 2)
        return self.quantity * (main + aux)

    def _calculate_price(self, pricelist: Pricelist) -> float:
        return self.surface_m2 * pricelist.pipe_fitting_metal_m2


class ElementFactory:
    """Creates the correct element object for a given row."""

    # Sort registered elements so more specific ones (subclasses)
    # are checked before less specific ones (BaseElement).
    # This relies on BaseElement being last in the MRO path.
    # We add a catch for BaseElement itself.
    _sorted_classes = sorted(
        [cls for cls in BaseElement.REGISTERED_ELEMENTS if cls != BaseElement],
        key=lambda x: -len(x.mro())
    )

    @staticmethod
    def create_element(row: dict, pricelist: dict) -> BaseElement:
        """Iterate registered classes and find the first one that can parse the row."""
        for ElementClass in ElementFactory._sorted_classes:
            if ElementClass.can_parse(row):
                return ElementClass(row, pricelist)
        
        # Fallback to BaseElement if no specific class matches
        return BaseElement(row, pricelist)