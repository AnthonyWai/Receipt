import json
import unittest
from unittest.mock import Mock, patch

from receipt_processor import API_URL, MODEL, extract_receipt_fields, is_unknown


class IsUnknownTests(unittest.TestCase):
    def test_recognizes_unknown_values(self):
        for value in ("", "unknown", " N/A ", None):
            with self.subTest(value=value):
                self.assertTrue(is_unknown(value))

    def test_keeps_known_values(self):
        self.assertFalse(is_unknown("Taste"))


class ExtractReceiptFieldsTests(unittest.TestCase):
    @patch("receipt_processor.requests.post")
    def test_extracts_fields_from_api_response(self, mock_post):
        fields = {
            "shop_name": "Taste",
            "category": "groceries",
            "date": "2026-10-27",
            "total_price": "125.50",
            "currency": "HKD",
        }
        response = Mock()
        response.json.return_value = {
            "choices": [{"message": {"content": json.dumps(fields)}}]
        }
        mock_post.return_value = response

        result = extract_receipt_fields("Taste\nTotal HKD 125.50", "test-api-key")

        self.assertEqual(result, fields)
        mock_post.assert_called_once()
        args, kwargs = mock_post.call_args
        self.assertEqual(args[0], API_URL)
        self.assertEqual(kwargs["json"]["model"], MODEL)
        self.assertIn("Taste\nTotal HKD 125.50", kwargs["json"]["messages"][1]["content"])
        self.assertEqual(kwargs["timeout"], 60)

    @patch("receipt_processor.requests.post")
    def test_normalizes_taste_name_from_ocr_text(self, mock_post):
        response = Mock()
        response.json.return_value = {
            "choices": [
                {
                    "message": {
                        "content": '{"shop_name": "ToSTe", "total_price": "125.50"}'
                    }
                }
            ]
        }
        mock_post.return_value = response

        result = extract_receipt_fields("ToSTe\nTotal HKD 125.50", "test-api-key")

        self.assertEqual(result["shop_name"], "Taste")


if __name__ == "__main__":
    unittest.main()