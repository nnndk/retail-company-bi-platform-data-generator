from src.excel_generator.excel_dimension_item import ExcelDimensionItem


class ExcelDimension:
    """
    This class is a representation of an OLAP-cube dimension.
    An instance of this class consists of fields (1 or more), time series, and facts (also 1 or more, one fact = [])
    """
    def __init__(self, dimension_name: str, fields: list[str], items: list[ExcelDimensionItem]):
        """
        Constructor. Each field from fields must be among items[i].values (if it's not -> error).
        All items[i].values should have equal keys and equal number of keys.
        :param dimension_name:
        :param fields: Fields of this cube dimension
        :param items: Values of each this cube dimension field
        """
        self.name = dimension_name
        self.fields = fields
        self.values = items
