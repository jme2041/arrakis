# TestArrakis.py: Python unit tests
# This requires win32com, part of pywin32 (pip install pywin32)

import unittest
import win32com.client
import pywintypes


class TestArrakeener(unittest.TestCase):
    '''Test the Arrakis.Arrakeener class'''

    def test_Operations(self):
        Paul = win32com.client.Dispatch('Arrakis.Arrakeener')

        Paul.FirstName = 'Paul'
        Paul.LastName = 'Atreides'
        Paul.Affiliation = 'House Atreides'
        Paul.Occupation = 'Kwisatz Haderach'

        self.assertEqual(Paul.FirstName, 'Paul')
        self.assertEqual(Paul.LastName, 'Atreides')
        self.assertEqual(Paul.Affiliation, 'House Atreides')
        self.assertEqual(Paul.Occupation, 'Kwisatz Haderach')

        energy = Paul.Energy
        solaris = Paul.Solaris
        spice = Paul.Spice

        self.assertGreaterEqual(energy, 1)
        self.assertLessEqual(energy, 100)

        self.assertGreaterEqual(solaris, 200000)
        self.assertLessEqual(solaris, 400000)

        self.assertEqual(spice, 0)

        # Try to sell without spice
        with self.assertRaisesRegex(pywintypes.com_error, 'Insufficient spice'):
            Paul.SellSpice()

        self.assertEqual(Paul.Energy, energy)
        self.assertEqual(Paul.Solaris, solaris)
        self.assertEqual(Paul.Spice, spice)

        # Try to eat without spice
        with self.assertRaisesRegex(pywintypes.com_error, 'Insufficient spice'):
            Paul.EatSpice()

        self.assertEqual(Paul.Energy, energy)
        self.assertEqual(Paul.Solaris, solaris)
        self.assertEqual(Paul.Spice, spice)

        # Mine some spice
        delta_spice = Paul.MineSpice()
        spice += delta_spice
        self.assertEqual(Paul.Spice, spice)

        # Energy and Solaris will have randomly decreased
        self.assertLess(Paul.Energy, energy)
        self.assertLess(Paul.Solaris, solaris)

        energy = Paul.Energy
        solaris = Paul.Solaris

        # Eat one tenth of our spice
        delta_spice = spice//10
        delta_energy = Paul.EatSpice(delta_spice)
        energy += delta_energy
        spice -= delta_spice

        self.assertEqual(Paul.Energy, energy)
        self.assertEqual(Paul.Solaris, solaris)
        self.assertEqual(Paul.Spice, spice)

        # Try to eat too much spice
        with self.assertRaisesRegex(pywintypes.com_error, 'Insufficient spice'):
            Paul.EatSpice(spice + 1)

        self.assertEqual(Paul.Energy, energy)
        self.assertEqual(Paul.Solaris, solaris)
        self.assertEqual(Paul.Spice, spice)

        # Sell half of our spice
        delta_spice = spice//2;
        delta_solaris = Paul.SellSpice(delta_spice)
        spice -= delta_spice
        solaris += delta_solaris

        self.assertEqual(Paul.Energy, energy)
        self.assertEqual(Paul.Solaris, solaris)
        self.assertEqual(Paul.Spice, spice)

        # Try to sell too much spice
        with self.assertRaisesRegex(pywintypes.com_error, 'Insufficient spice'):
            Paul.SellSpice(spice + 1)

        self.assertEqual(Paul.Energy, energy)
        self.assertEqual(Paul.Solaris, solaris)
        self.assertEqual(Paul.Spice, spice)

        # Check for integer overflow
        with self.assertRaisesRegex(pywintypes.com_error, 'Integer overflow'):
            Paul.MineSpice((2**63)-2)

        self.assertEqual(Paul.Energy, energy)
        self.assertEqual(Paul.Solaris, solaris)
        self.assertEqual(Paul.Spice, spice)

    def test_Clone(self):
        Duncan = win32com.client.Dispatch('Arrakis.Arrakeener')

        Duncan.FirstName = 'Duncan'
        Duncan.LastName = 'Idaho'
        Duncan.Affiliation = 'House Atreides'
        Duncan.Occupation = 'Swordmaster'

        # Create an alias of Duncan
        Alias = Duncan

        # Create a copy of Duncan
        Ghola = Duncan.Clone()
        Ghola.Occupation = 'Ghola'

        # Check object identity
        self.assertEqual(id(Duncan), id(Alias))
        self.assertNotEqual(id(Duncan), id(Ghola))

        # Compare data between Duncan and his alias
        self.assertEqual(Duncan.FirstName, Alias.FirstName)
        self.assertEqual(Duncan.LastName, Alias.LastName)
        self.assertEqual(Duncan.Affiliation, Alias.Affiliation)
        self.assertEqual(Duncan.Occupation, Alias.Occupation)

        # Compare data between Duncan and his Ghola
        self.assertEqual(Duncan.FirstName, Ghola.FirstName)
        self.assertEqual(Duncan.LastName, Ghola.LastName)
        self.assertEqual(Duncan.Affiliation, Ghola.Affiliation)
        self.assertNotEqual(Duncan.Occupation, Ghola.Occupation)

        # Switching an attribute in Duncan's alias switches it in Duncan but not the Ghola
        Alias.Occupation = 'Ambassador to the Fremen'
        self.assertEqual(Duncan.Occupation, 'Ambassador to the Fremen')
        self.assertEqual(Ghola.Occupation, 'Ghola')
