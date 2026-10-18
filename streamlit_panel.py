"""Standalone demo panel; import and call render_queue_panel() from app.py.
No salary information, PDF bytes or real phone numbers are transmitted.
"""
import streamlit as st
from queue_client import client, enqueue_batch, batch_jobs


def render_queue_panel(employees=None):
    st.subheader('Gửi phiếu lương Zalo — THỬ NGHIỆM (KHÔNG GỬI THẬT)')
    st.warning('Chỉ mô phỏng hàng đợi. Không kết nối hay gửi tin nhắn Zalo.')
    # Sample data. Replace with sanitized employee list from app.py after validating mappings.
    if employees is None:
        employees = [
            {'employee_code': 'NV001', 'employee_label': 'Nhân viên thử nghiệm A', 'recipient_alias': 'zalo_test_a', 'pdf_filename': 'phieu_luong_NV001.pdf'},
            {'employee_code': 'NV002', 'employee_label': 'Nhân viên thử nghiệm B', 'recipient_alias': 'zalo_test_b', 'pdf_filename': 'phieu_luong_NV002.pdf'},
        ]
    lookup = {e['employee_code']: e for e in employees}
    chosen = st.multiselect('Nhân viên đưa vào hàng đợi', list(lookup), default=list(lookup))
    period = st.text_input('Kỳ lương', '10/2026')
    confirm = st.checkbox('Tôi xác nhận đây là chế độ thử nghiệm, không gửi Zalo', key='zalo_dryrun_confirm')
    if st.button('Tạo hàng đợi thử nghiệm', disabled=not chosen or not confirm, key='zalo_dryrun_enqueue'):
        try:
            db = client(st.secrets['supabase']['url'], st.secrets['supabase']['service_role_key'])
            batch_id = enqueue_batch(db, period, [lookup[code] for code in chosen])
            st.session_state['zalo_dryrun_batch_id'] = batch_id
            st.success(f'Đã tạo {len(chosen)} yêu cầu thử nghiệm. Batch: {batch_id}')
        except Exception as exc:
            st.error(f'Không tạo được hàng đợi: {type(exc).__name__}. Kiểm tra cấu hình và kết nối.')
    batch_id = st.session_state.get('zalo_dryrun_batch_id')
    if batch_id:
        st.caption(f'Batch ID: {batch_id}')
        if st.button('Làm mới trạng thái', key='zalo_dryrun_refresh'):
            pass
        try:
            db = client(st.secrets['supabase']['url'], st.secrets['supabase']['service_role_key'])
            st.dataframe(batch_jobs(db, batch_id), use_container_width=True)
        except Exception:
            st.error('Không tải được trạng thái hàng đợi.')

