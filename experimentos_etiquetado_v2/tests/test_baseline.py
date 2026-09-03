import unittest

from experimentos_etiquetado_v2.src.baseline import percentage


class BaselineTests(unittest.TestCase):
    def test_percentage(self) -> None:
        self.assertEqual(percentage(1, 4), 25.0)

    def test_percentage_with_empty_total(self) -> None:
        self.assertEqual(percentage(1, 0), 0.0)


if __name__ == "__main__":
    unittest.main()
