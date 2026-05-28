from app.ai_parser import fallback_parse
from app.validator import validate_action_plan


def _model_map():
    return {
        "settings": {"default_outputs": ["project_npv", "project_irr", "equity_npv", "equity_irr"]},
        "inputs": {
            "investment_cost_change": {"aliases": ["tong muc dau tu", "tmdt", "chi phi dau tu"], "editable": True, "min": -0.3, "max": 1.0},
            "loan_interest_rate": {"aliases": ["lai vay", "lai suat", "lai ngan hang", "interest rate"], "editable": True, "min": 0, "max": 0.25},
            "selling_price_change": {"aliases": ["gia ban", "doanh thu"], "editable": True, "min": -0.5, "max": 1.0},
            "vat_rate": {"aliases": ["vat", "thue vat"], "editable": True, "min": 0, "max": 0.15},
            "cit_rate": {"aliases": ["tndn", "cit", "thue tndn"], "editable": True, "min": 0, "max": 0.35},
            "marketing_rate": {"aliases": ["marketing", "truyen thong"], "editable": True, "min": 0, "max": 0.2},
            "premium_change": {"aliases": ["premium", "premium giao dich"], "editable": True, "min": -1.0, "max": 1.0, "base_value": 260},
        },
        "outputs": {
            "project_npv": {"aliases": ["npv du an", "npv"]},
            "project_irr": {"aliases": ["irr du an", "irr"]},
            "equity_npv": {"aliases": ["npv chu dau tu"]},
            "equity_irr": {"aliases": ["irr chu dau tu"]},
        },
    }


def _changes(command):
    plan = fallback_parse(command, _model_map())
    return {c.parameter: float(c.value) for c in plan.changes}


def test_tmdt_from_to():
    changes = _changes("Tăng tổng mức đầu tư từ 1000 tỷ lên 1200 tỷ")
    assert round(changes["investment_cost_change"], 6) == 0.2


def test_interest_rate():
    changes = _changes("Tăng lãi vay lên 8%")
    assert round(changes["loan_interest_rate"], 6) == 0.08


def test_interest_rate_without_percent_sign():
    changes = _changes("Cho lãi ngân hàng là 8")
    assert round(changes["loan_interest_rate"], 6) == 0.08


def test_price_decrease():
    changes = _changes("Giảm giá bán 5%")
    assert round(changes["selling_price_change"], 6) == -0.05


def test_revenue_decrease_alias():
    changes = _changes("Nếu doanh thu giảm 7% thì sao")
    assert round(changes["selling_price_change"], 6) == -0.07


def test_multi_scenario():
    changes = _changes("Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%")
    assert round(changes["investment_cost_change"], 6) == 0.15
    assert round(changes["selling_price_change"], 6) == -0.05
    assert round(changes["loan_interest_rate"], 6) == 0.08


def test_tax_rates_with_no_comma():
    changes = _changes("VAT 10% TNDN 20%")
    assert round(changes["vat_rate"], 6) == 0.10
    assert round(changes["cit_rate"], 6) == 0.20


def test_output_only_command_allowed_by_validator():
    model_map = _model_map()
    plan = fallback_parse("Cho tôi NPV và IRR hiện tại", model_map)
    assert plan.changes == []
    assert "project_npv" in plan.requested_outputs
    assert "project_irr" in plan.requested_outputs
    validate_action_plan(plan, model_map)


def test_premium_change_percent():
    changes = _changes("Premium giao dich tang 5%")
    assert round(changes["premium_change"], 6) == 0.05


def test_premium_change_absolute_target_with_base_value():
    changes = _changes("Premium giao dich la 270")
    assert round(changes["premium_change"], 6) == round(270 / 260 - 1, 6)
