
#use dataclass 
from dataclasses import dataclass


@dataclass(frozen=True)
class HourRecord:
    record_code: str
    record_name: str
    rt:str
    t15:str
    r20:str

@dataclass(frozen=True)
class Employee:
    employee_id: str
    name: str
    special_code: str
    hour_records: list[HourRecord]

@dataclass(frozen=True)
class EmployeePage:
    employees: list[Employee]