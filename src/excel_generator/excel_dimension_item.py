class ExcelDimensionItem:
    """
    Instances fo this class store a dimension entry (id and all dimension field values)
    """
    def __init__(self, item_id: int, values: dict[str, str]):
        """
        Constructor. 'values' keys are field names, its values are field values.
        Example: {'Item name': 'Pen', 'Price': '50'}
        :param item_id: Dimension entry id (for instance, FK in relational table)
        :param values: Dict of fields and its values for this entry
        """
        self.id = item_id
        self.values = values
