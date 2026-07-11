"""
test_customer_consistency.py — tests for the customer-mismatch safety net
(data_pipeline.check_customer_consistency).
"""
import pandas as pd
import pytest

from data_pipeline import check_customer_consistency, _extract_customers


def _det(customer):
    return pd.DataFrame({'Name': ['HOST01'], 'Vulnerability Name': ['CVE-2025-0001'],
                         'Customer': [customer]})


def _dev(customer):
    return pd.DataFrame({'Device': ['HOST01'], 'Last Response': ['01/01/2026'],
                         'Customer Name': [customer]})


def test_matching_customers_pass():
    check_customer_consistency({'Detections export': _det('Acme Corp'),
                                'Device report':     _dev('Acme Corp')})


def test_case_and_whitespace_insensitive():
    check_customer_consistency({'Detections export': _det('  ACME corp '),
                                'Device report':     _dev('Acme Corp')})


def test_mismatch_raises_with_both_files_named():
    with pytest.raises(ValueError) as exc:
        check_customer_consistency({'Detections export': _det('Acme Corp'),
                                    'Device report':     _dev('Beta Ltd')})
    msg = str(exc.value)
    assert 'Acme Corp' in msg and 'Beta Ltd' in msg
    assert 'Detections export' in msg and 'Device report' in msg


def test_previous_report_mismatch_raises():
    prev = pd.DataFrame({'Name': ['HOST02'], 'Vulnerability Name': ['CVE-2025-0002'],
                         'Customer': ['Other Client']})
    with pytest.raises(ValueError):
        check_customer_consistency({'Detections export (current)': _det('Acme Corp'),
                                    'Previous report': prev})


def test_file_without_customer_column_is_skipped():
    dev_no_cust = pd.DataFrame({'Device': ['HOST01'], 'Last Response': ['01/01/2026']})
    check_customer_consistency({'Detections export': _det('Acme Corp'),
                                'Device report':     dev_no_cust})


def test_blank_customer_values_are_skipped():
    check_customer_consistency({'Detections export': _det('Acme Corp'),
                                'Device report':     _dev('')})
    check_customer_consistency({'Detections export': _det('Acme Corp'),
                                'Device report':     _dev('nan')})


def test_none_frame_is_skipped():
    check_customer_consistency({'Detections export': _det('Acme Corp'),
                                'Device report':     None})


def test_multiple_customers_in_one_file_raises():
    mixed = pd.DataFrame({'Name': ['H1', 'H2'],
                          'Vulnerability Name': ['CVE-2025-1', 'CVE-2025-2'],
                          'Customer': ['Acme Corp', 'Beta Ltd']})
    with pytest.raises(ValueError):
        check_customer_consistency({'Detections export': mixed})


def test_extract_customers_prefers_customer_over_customer_name():
    df = pd.DataFrame({'Customer': ['Acme Corp'], 'Customer Name': ['Ignored']})
    assert list(_extract_customers(df).values()) == ['Acme Corp']


def test_extract_customers_reads_customer_name_variant():
    assert list(_extract_customers(_dev('Acme Corp')).values()) == ['Acme Corp']