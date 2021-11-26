import re
from datetime import datetime, timedelta
from SlyAPI import *

RE_A1 = re.compile(
    # like: 'page_name'!A1:B2
    r"(?:'(?P<page_name>\w+)'!)?" # TODO: is \w exactly what sheets allows?
    # can be just start_col (e.g. 'A'), or just cell (e.g. 'A1')
    r"(?P<start_col>[a-z]{1,2})(?P<start_row>\d+)?" 
    # ranges (e.g. 'A1:B2') or column ranges (e.g. 'A:B')
    r"(?::(?P<end_col>[a-z]{1,2})(?P<end_row>\d+)?)?", 
    re.IGNORECASE
)

ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
LOTUS123_EPOCH = datetime(1899, 12, 30)

class Sheet(WebAPI):
    """
    Class for handling sheets
    """
    base_url = "https://sheets.googleapis.com/v4/spreadsheets"
    id: str
    n_columns: int|None
    headers: list[str]|None

    def __init__(self, sheet_id: str, auth: OAuth2User):
        super().__init__(auth=auth)
        self.id = sheet_id

    async def properties(self) -> dict[str, Any]:
        """
        Get sheet metadata
        """
        return await self.get_json(
            F"{self.base_url}/{self.id}/properties",
        )

    