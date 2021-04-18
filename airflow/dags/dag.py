import datetime
from airflow import DAG
from airflow.operators.python import PythonOperator
from airflow.utils.dates import days_ago

import sys
sys.path.insert(1, '/Users/pathairs/Documents/projects/automate_stuff_with_pyautogui/airflow/scripts')
from login_vpn import main as login_vpn
from run import extract_zfiaraging_report

import json

default_args = {
    'owner' : 'pathairs'
}

dag = DAG(
    'sap_report_extraction_pipeline',
    default_args = default_args,
    description = 'to extract ZFIARAGING report from SAP GUI application daily',
    schedule_interval = None,
    start_date = days_ago(1)
)

login_vpn_task = PythonOperator(
    task_id = 'login_vpn',
    dag = dag,
    python_callable=login_vpn
)

with dag:
    with open("airflow/scripts/input.json") as json_file:
        data = json.load(json_file)        
        for bu, inputs  in data.items():
            task = PythonOperator(
                task_id = f'extract_report_{inputs["bu_name"]}_{inputs["profit_center"]}',
                python_callable=extract_zfiaraging_report,
                op_kwargs={
                    **inputs,
                    'aging_date' : days_ago(1).strftime("%d.%m.%Y"),
                    'export_date' : days_ago(1).strftime("%Y%m%d"),
                    'sap_server' : 'dev',
                    'config' : None
                }
            )