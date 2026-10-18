"""Server-side Supabase queue access. Never import this in browser JS."""
import os
from supabase import create_client


def client(url=None, key=None):
    url = url or os.environ.get('SUPABASE_URL')
    key = key or os.environ.get('SUPABASE_SERVICE_ROLE_KEY')
    if not url or not key:
        raise RuntimeError('Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY')
    return create_client(url, key)


def enqueue_batch(db, period, employees):
    """employees: [{employee_code, employee_label, recipient_alias, pdf_filename}].
    Dry-run only. Do not pass salary values, PDFs, or phone numbers.
    """
    if not employees:
        raise ValueError('No employees selected')
    if len(employees) > 500:
        raise ValueError('Maximum 500 employees per batch')
    codes = [e['employee_code'] for e in employees]
    if len(codes) != len(set(codes)):
        raise ValueError('Duplicate employee codes')
    batch = db.table('payroll_batches').insert({'period': period, 'dry_run': True}).execute().data[0]
    rows = [{**e, 'batch_id': batch['id']} for e in employees]
    try:
        db.table('payroll_jobs').insert(rows).execute()
    except Exception:
        db.table('payroll_batches').delete().eq('id', batch['id']).execute()
        raise
    return batch['id']


def batch_jobs(db, batch_id):
    return db.table('payroll_jobs').select('employee_code,employee_label,recipient_alias,pdf_filename,status,result_note,attempt_count').eq('batch_id',batch_id).order('employee_code').execute().data

