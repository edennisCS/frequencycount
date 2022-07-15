import unittest

from app import counted_values


class TestProblem(unittest.TestCase):
    def setUp(self) -> None:
        self.dictionary = {'Queen':  {
  'Word(Total Occurrences)': 68,
  'Documents': [
    'alice_in_wonderland'
  ],
  'Sentences': [
    'An\ninvitation from the *Queen* to play croquet  '
  ]
}}
        self.dictionary_fail = {'King': {
            'Word(Total Occurrences)': 10,
            'Documents': [
                'alice_in_wonderland'
            ],
            'Sentences': [
                'An\ninvitation from the *King* to play croquet  '
            ]
        }}

    def test_counted(self):
        self.assertEqual(counted_values(self.dictionary), {'Queen': {'Documents': ['alice_in_wonderland'],
           'Sentences': ['An\ninvitation from the *Queen* to play croquet  '],
           'Word(Total Occurrences)': 68}})
        self.assertEqual(counted_values(self.dictionary_fail), {})

if __name__ == '__main__':
    unittest.main()
