import unittest
from pathlib import Path

from app.template_writer import prepare_variation_rows
from app.template_validator import _validate_variation_titles


class VariationTitleRulesTest(unittest.TestCase):
    def test_prepare_variation_rows_generates_parent_and_child_titles(self):
        rows = [
            {
                "product_name": "Floral Card Holder Picks",
                "route": "Haul Generic Variation",
                "sku": "FLORAL-GOLD",
                "color": "Gold",
                "set_count": "50pcs 2set",
                "title": "Floral Card Holder Picks, Flower Bouquet Card Sticks for Gift Cards and Table Centerpieces, Gold",
            },
            {
                "product_name": "Floral Card Holder Picks",
                "route": "Haul Generic Variation",
                "sku": "FLORAL-SILVER",
                "color": "Silver",
                "set_count": "50pcs 2set",
                "title": "Floral Card Holder Picks, Flower Bouquet Card Sticks for Gift Cards and Table Centerpieces, Silver",
            },
        ]

        prepared = prepare_variation_rows(rows, Path("/tmp/floral_card_holder"))

        parent = prepared[0]
        children = prepared[1:]
        self.assertEqual(parent["parentage_level"], "Parent")
        self.assertTrue(parent["title"].startswith("50 Pcs 2 Set "))
        self.assertIn("Multiple Colors Available", parent["title"])
        self.assertNotIn("Gold", parent["title"])
        self.assertLessEqual(len(parent["title"]), 125)

        self.assertTrue(children[0]["title"].startswith("50 Pcs 2 Set "))
        self.assertIn("Gold", children[0]["title"])
        self.assertNotIn("Multiple Colors Available", children[0]["title"])
        self.assertLessEqual(len(children[0]["title"]), 125)

    def test_variation_title_self_check_flags_parent_and_child_mismatches(self):
        row_infos = {
            7: {
                "row": 7,
                "sku": "PARENT",
                "parentage": "Parent",
                "parent_sku": "",
                "variation_theme": "Color",
                "title": "Floral Card Holder Picks, Gold",
                "color": "",
                "size": "",
            },
            8: {
                "row": 8,
                "sku": "CHILD-GOLD",
                "parentage": "Child",
                "parent_sku": "PARENT",
                "variation_theme": "Color",
                "title": "Floral Card Holder Picks, Multiple Colors Available",
                "color": "Gold",
                "size": "",
            },
            9: {
                "row": 9,
                "sku": "CHILD-SILVER",
                "parentage": "Child",
                "parent_sku": "PARENT",
                "variation_theme": "Color",
                "title": "Floral Card Holder Picks, Silver",
                "color": "Silver",
                "size": "",
            },
        }

        findings = []
        _validate_variation_titles(findings, row_infos)
        messages = "\n".join(item["message"] for item in findings)

        self.assertIn("多颜色 Parent", messages)
        self.assertIn("Child 行，但标题包含父体总结词", messages)
        self.assertIn("颜色变体 Child，但标题未包含自身颜色", messages)


if __name__ == "__main__":
    unittest.main()
