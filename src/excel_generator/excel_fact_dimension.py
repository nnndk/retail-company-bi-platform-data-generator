from src.tools.common_tools import *


class ExcelFactDimension:
    def __init__(self, fact_name: str, min_value: int, max_value: int):
        self.name = fact_name
        self.min_value = min_value
        self.max_value = max_value

    def generate_fact_value(self) -> int:
        return get_random_int(self.min_value, self.max_value)
