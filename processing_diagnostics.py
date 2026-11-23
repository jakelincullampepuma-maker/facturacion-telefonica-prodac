"""Tiempos por fase; no registra documentos, teléfonos ni credenciales."""
from time import perf_counter


class PhaseTimings:
    def __init__(self):
        self.values = dict.fromkeys(('pdf_read_ms', 'excel_read_ms', 'inventory_match_ms',
                                    'reconciliation_ms', 'workbook_ms'), 0.0)
        self.last = perf_counter()

    def mark(self, phase):
        now = perf_counter()
        self.values[phase] += (now - self.last) * 1000
        self.last = now

    def run(self, phase, function, *args, **kwargs):
        start = perf_counter()
        try:
            return function(*args, **kwargs)
        finally:
            self.values[phase] += (perf_counter() - start) * 1000
            self.last = perf_counter()

    def snapshot(self):
        return {key: round(value, 3) for key, value in self.values.items()}
