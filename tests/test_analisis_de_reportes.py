import unittest

from scripts.analysis.analisis_de_reportes import infer_type_and_family, is_family_noise


class LabelInferenceTests(unittest.TestCase):
    def test_pdf_heuristic_tokens_are_not_used_as_families(self) -> None:
        noisy_family_tokens = {
            "malurl",
            "static",
            "badfile",
            "obfuscated",
            "shellcode",
            "iframe",
            "bomb",
            "camelot",
            "gerphish",
            "rbloxphish",
            "filerepmalware",
        }

        for token in noisy_family_tokens:
            with self.subTest(token=token):
                self.assertTrue(is_family_noise(token))

    def test_noise_tokens_do_not_beat_real_family_votes(self) -> None:
        detections = [
            {
                "engine": "BitDefender",
                "result": "Static AI - Suspicious PDF",
                "engine_weight": 1.8,
            },
            {
                "engine": "Microsoft",
                "result": "Trojan:Win32/Salgorea.VRR!MTB",
                "engine_weight": 2.2,
            },
            {
                "engine": "Kaspersky",
                "result": "Backdoor.Win32.Salgorea.ie",
                "engine_weight": 2.0,
            },
        ]

        inferred = infer_type_and_family(detections)

        self.assertEqual(inferred["familia_probable"], "salgorea")
        self.assertIn(inferred["tipo_probable"], {"trojan", "backdoor"})

    def test_all_noise_family_candidates_fall_back_to_uninferred(self) -> None:
        detections = [
            {
                "engine": "BitDefender",
                "result": "Phishing/PDF.Malurl.XG4",
                "engine_weight": 1.8,
            },
            {
                "engine": "SentinelOne",
                "result": "Static AI - Suspicious PDF",
                "engine_weight": 1.3,
            },
            {
                "engine": "Malwarebytes",
                "result": "BehavesLike.PDF.BadFile.cb",
                "engine_weight": 1.4,
            },
        ]

        inferred = infer_type_and_family(detections)

        self.assertEqual(inferred["familia_probable"], "sin_inferir")
        self.assertEqual(inferred["tipo_probable"], "phishing")


if __name__ == "__main__":
    unittest.main()
