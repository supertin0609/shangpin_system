import unittest

from app.listing_rules import validate_listing_row


def _title_findings(title):
    return [
        message
        for field, message, _fix in validate_listing_row({"title": title, "material": ""})
        if field in {"title", "copy"}
    ]


class ListingRulesTest(unittest.TestCase):
    def test_safe_pet_toy_title_passes_title_rules(self):
        title = (
            "Cat Chew Toy, Interactive Bite Toy for Indoor Cats and Kittens, "
            "Multiple Shapes Available, Active Play and Chewing Activity"
        )

        self.assertEqual(_title_findings(title), [])

    def test_pet_dental_and_flavor_claims_are_blocked_in_title(self):
        title = (
            "Fish Shape Cat Chew Toy, Natural Mint Flavor Dental Toy for Indoor Cats, "
            "Kittens, Teeth Cleaning and Chewing Activity"
        )

        messages = "\n".join(_title_findings(title))

        self.assertIn("医疗、护理、清洁、防护或结果型功效表达", messages)
        self.assertIn("摄入暗示或绝对化成分宣称", messages)

    def test_forbidden_characters_and_promotional_terms_are_blocked_in_title(self):
        title = (
            "Travel Storage Bag! Free Shipping, Organizer Case for Home, Closet, "
            "Bedroom, Moving, Sorting and Daily Use"
        )

        messages = "\n".join(_title_findings(title))

        self.assertIn("Amazon 禁用特殊字符", messages)
        self.assertIn("促销、平台背书、物流或售后类表达", messages)

    def test_non_risky_verifiable_material_can_appear_in_title(self):
        title = (
            "Velvet Pet Collar with Bell, Adjustable Collar for Cats and Small Dogs, "
            "Daily Wear and Indoor Outdoor Use, Dark Brown"
        )

        self.assertEqual(_title_findings(title), [])

    def test_ordinary_material_terms_can_appear_in_title(self):
        title = (
            "Plastic Pet Collar with Bell, Adjustable Collar for Cats and Small Dogs, "
            "Daily Wear and Indoor Outdoor Use, Dark Brown"
        )

        self.assertEqual(_title_findings(title), [])

    def test_wooden_handle_can_appear_without_absolute_material_claims(self):
        title = (
            "Coffee Grinder Cleaning Brush with Wooden Handle, Espresso Machine "
            "Barista Tool for Home Kitchen Coffee Bar"
        )

        self.assertEqual(_title_findings(title), [])

    def test_natural_material_claim_is_blocked_in_title(self):
        title = (
            "Coffee Grinder Cleaning Brush with Wooden Handle, Natural Bristles, "
            "Espresso Machine Barista Tool for Home Kitchen"
        )

        messages = "\n".join(_title_findings(title))

        self.assertIn("摄入暗示或绝对化成分宣称", messages)

    def test_sensitive_material_claims_are_blocked_in_title(self):
        title = (
            "Food Grade Silicone Pet Bowl, Travel Bowl for Cats and Small Dogs, "
            "Daily Feeding, Home Use and Outdoor Trips"
        )

        messages = "\n".join(_title_findings(title))

        self.assertIn("敏感材质或合规宣称", messages)


if __name__ == "__main__":
    unittest.main()
