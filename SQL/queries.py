"""
queries.py

SQL Queries used in other scripts
"""

# Imports
import psycopg2
import sqlalchemy
import os


save_directory = (
    r"C:\Users\Mitchell.Hollberg\Community Foundation for Greater Atlanta\IT - Documents\Public\Reports"
)

POSTGRES_HOST = '10.20.141.105'
POSTGRES_USER = 'postgres'
CSUITE_POSTGRES_DB = 'csuite'
POSTGRES_PASSWORD = os.environ['CSUITE_POSTGRES_SQL_SERVER_PWD']
# POSTGRES_PASSWORD = os.environ['POSTGRES_DW_PASSWORD']

conn = psycopg2.connect(host=POSTGRES_HOST,
                        database=CSUITE_POSTGRES_DB,
                        user=POSTGRES_USER,
                        password=POSTGRES_PASSWORD)

cursor = conn.cursor()

# Create sqlalchemy connection
engine = sqlalchemy.create_engine(f'postgresql://{POSTGRES_USER}'
                                  f':{POSTGRES_PASSWORD}@{POSTGRES_HOST}/'
                                  f'{CSUITE_POSTGRES_DB}')


QUERIES = {


    'open_daf_funds': r"""
    /*
    Selects all open DAF Funds 
    */
    select
        short_name,
        funits.funit_id,
        funits."name" as "Fund Name",
        employees."name" as "Steward Name",
        funits.created_ts::date as "Created Date",
        close_ts::date as "Close Date",
        fgroups."name" as "Fund Group",
        current_fundbalance
    
    from
        public.funits as funits
        
    LEFT JOIN LATERAL 
        json_array_elements(funits.custom_fields::json) AS j(custom_field) on true
    
    left join fgroups 
        on fgroups.fgroup_id = funits.fgroup_id 
    
    left join employees_v employees
     on employees.employee_id = funits.steward_employee_id 
    
    where funits.close_ts is NULL and
    funits.fgroup_id = 1039       -- {1039: Donor Advised}
        
    group by 
        funits.short_name,
        funits.funit_id,
        funits.name,
        funits.steward_employee_id,
        employees.name,
        funits.created_ts,
        close_ts,
        funits.fgroup_id,
        fgroups."name",
        funits.funit_type_id,
        subof_funit_id,
        current_fundbalance,
        current_spendable,
        current_principal
        
    order by funits.short_name asc, funits."name" asc;
    """,

    # ******************
    'fund_files': """
    /*
    Return all files attached to "Fund" objects in CSuite
    */
    select 
        file_id,
        files.name as "file_name",
        files.created_ts,
        files.employee_id,
        employees.name as "Employee Name",
        ref,
        ref_id as funit_id,     -- Rename to align standardize
        description,
        content_type,
        ref_name,
        sticky,
        files.file_cat_id,
        file_cat.file_cat_name as "File Category",
        filesize,
        filesize_mb,
        shared_file
        
        from files_v_r files
        
        left join file_categories_v_r file_cat
        on file_cat.file_cat_id = files.file_cat_id
        
        left join employees_v employees
        on employees.employee_id = files.employee_id
        
        where ref = 'funit'

"""

}
