from kuchenland_importer.infrastructure.excel.formula_routes import rebind_formula


def test_retarget_quoted_and_unquoted_sheet_references():
    assert rebind_formula("=Весна!RC1+'Весна'!R2C3", {"Весна": "Новый сезон"}) == (
        "='Новый сезон'!RC1+'Новый сезон'!R2C3"
    )


def test_formula_text_and_similarly_named_sheets_are_not_rewritten():
    formula = '=IF(RC1="Весна!",ПоздняяВесна!RC1,Весна!RC1)'
    assert rebind_formula(formula, {"Весна": "Осень"}) == (
        "=IF(RC1=\"Весна!\",ПоздняяВесна!RC1,'Осень'!RC1)"
    )


def test_escaped_apostrophes_in_worksheet_names():
    assert rebind_formula("='Supplier''s'!RC1", {"Supplier's": "Season's"}) == ("='Season''s'!RC1")
