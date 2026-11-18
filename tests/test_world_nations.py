import unittest

from nations import WORLD_CHAMPIONSHIP_48, flag_file, nation_code


class WorldNationTests(unittest.TestCase):
    def test_all_48_confirmed_teams_have_flags(self):
        self.assertEqual(len(WORLD_CHAMPIONSHIP_48), 48)
        self.assertEqual(len(set(WORLD_CHAMPIONSHIP_48)), 48)
        for label in WORLD_CHAMPIONSHIP_48:
            with self.subTest(label=label):
                self.assertIsNotNone(nation_code(label))
                self.assertIsNotNone(flag_file(label))


if __name__ == "__main__":
    unittest.main()
