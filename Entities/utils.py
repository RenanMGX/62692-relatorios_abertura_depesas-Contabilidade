from datetime import datetime
from dateutil.relativedelta import relativedelta

def ultimo_dia_mes(date:datetime) -> datetime:
    return ((date.replace(day=1) + relativedelta(months=1)) - relativedelta(days=1)).replace(hour=23, minute=59, second=59, microsecond=999999)

def primeiro_dia_mes(date: datetime) -> datetime:
    return date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    
if __name__ == "__main__":
    pass