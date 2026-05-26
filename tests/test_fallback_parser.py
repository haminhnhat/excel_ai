from app.ai_parser import fallback_parse


def test_tmdt_from_to():
    model_map = {
        "settings": {"default_outputs": ["project_npv"]},
        "outputs": {"project_npv": {}},
    }
    plan = fallback_parse("Tăng tổng mức đầu tư từ 1000 tỷ lên 1200 tỷ", model_map)
    assert plan.changes[0].parameter == "investment_cost_change"
    assert round(float(plan.changes[0].value), 6) == 0.2


def test_interest_rate():
    model_map = {"settings": {"default_outputs": []}, "outputs": {}}
    plan = fallback_parse("Tăng lãi vay lên 8%", model_map)
    assert plan.changes[0].parameter == "loan_interest_rate"
    assert round(float(plan.changes[0].value), 6) == 0.08
