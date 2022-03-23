// TestArrakis.cs: C# unit tests
// Add the Arrakis type library to the C# project as a COM reference

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Arrakis;

namespace TestArrakisCS
{
    [TestClass]
    public class TestArrakeener
    {
        [TestMethod]
        public void Operations()
        {
            Arrakeener Paul = new()
            {
                FirstName = "Paul",
                LastName = "Atreides",
                Affiliation = "House Atreides",
                Occupation = "Kwisatz Haderach"
            };

            Assert.AreEqual("Paul", Paul.FirstName);
            Assert.AreEqual("Atreides", Paul.LastName);
            Assert.AreEqual("House Atreides", Paul.Affiliation);
            Assert.AreEqual("Kwisatz Haderach", Paul.Occupation);

            long energy = Paul.Energy;
            long solaris = Paul.Solaris;
            long spice = Paul.Spice;

            Assert.IsTrue(energy >= 1);
            Assert.IsTrue(energy <= 100);

            Assert.IsTrue(solaris >= 200000);
            Assert.IsTrue(solaris <= 400000);

            Assert.AreEqual(0, spice);

            // Try to sell without spice
            try
            {
                Paul.EatSpice();
                Assert.Fail("Exception not thrown");
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                StringAssert.Contains(e.Message, "Insufficient spice");
            }
            catch
            {
                Assert.Fail("The wrong exception was thrown");
            }

            Assert.AreEqual(energy, Paul.Energy);
            Assert.AreEqual(solaris, Paul.Solaris);
            Assert.AreEqual(spice, Paul.Spice);

            // Try to eat without spice
            try
            {
                Paul.EatSpice();
                Assert.Fail("Exception not thrown");
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                StringAssert.Contains(e.Message, "Insufficient spice");
            }
            catch
            {
                Assert.Fail("The wrong exception was thrown");
            }

            Assert.AreEqual(energy, Paul.Energy);
            Assert.AreEqual(solaris, Paul.Solaris);
            Assert.AreEqual(spice, Paul.Spice);

            // Mine some spice
            long delta_spice = Paul.MineSpice();
            spice += delta_spice;
            Assert.AreEqual(spice, Paul.Spice);

            // Energy and solaris will have randomly decreased
            Assert.IsTrue(Paul.Energy < energy);
            Assert.IsTrue(Paul.Solaris < solaris);

            energy = Paul.Energy;
            solaris = Paul.Solaris;

            // Eat one tenth of our spice
            delta_spice = spice / 10;
            long delta_energy = Paul.EatSpice(delta_spice);
            energy += delta_energy;
            spice -= delta_spice;

            Assert.AreEqual(energy, Paul.Energy);
            Assert.AreEqual(solaris, Paul.Solaris);
            Assert.AreEqual(spice, Paul.Spice);

            // Try to eat too much spice
            try
            {
                Paul.EatSpice(spice + 1);
                Assert.Fail("Exception not thrown");
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                StringAssert.Contains(e.Message, "Insufficient spice");
            }
            catch
            {
                Assert.Fail("The wrong exception was thrown");
            }

            Assert.AreEqual(energy, Paul.Energy);
            Assert.AreEqual(solaris, Paul.Solaris);
            Assert.AreEqual(spice, Paul.Spice);

            // Sell half of our spice
            delta_spice = spice / 2;
            long delta_solaris = Paul.SellSpice(delta_spice);
            spice -= delta_spice;
            solaris += delta_solaris;

            Assert.AreEqual(energy, Paul.Energy);
            Assert.AreEqual(solaris, Paul.Solaris);
            Assert.AreEqual(spice, Paul.Spice);

            // Try to sell too much spice
            try
            {
                Paul.SellSpice(spice + 1);
                Assert.Fail("Exception not thrown");
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                StringAssert.Contains(e.Message, "Insufficient spice");
            }
            catch
            {
                Assert.Fail("The wrong exception was thrown");
            }

            Assert.AreEqual(energy, Paul.Energy);
            Assert.AreEqual(solaris, Paul.Solaris);
            Assert.AreEqual(spice, Paul.Spice);

            // Check for integer overflow
            try
            {
                Paul.MineSpice(System.Int64.MaxValue - 1);
                Assert.Fail("Exception not thrown");
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                StringAssert.Contains(e.Message, "Integer overflow");
            }
            catch
            {
                Assert.Fail("The wrong exception was thrown");
            }

            Assert.AreEqual(energy, Paul.Energy);
            Assert.AreEqual(solaris, Paul.Solaris);
            Assert.AreEqual(spice, Paul.Spice);
        }

        [TestMethod]
        public void Clone()
        {
            Arrakeener Duncan = new()
            {
                FirstName = "Duncan",
                LastName = "Idaho",
                Affiliation = "House Atreides",
                Occupation = "Swordmaster"
            };

            // Create an alias of Duncan
            Arrakeener Alias = Duncan;

            // Create a copy of Duncan
            Arrakeener Ghola = Duncan.Clone();
            Ghola.Occupation = "Ghola";

            // Compare data between Duncan and his alias
            Assert.AreEqual(Duncan.FirstName, Alias.FirstName);
            Assert.AreEqual(Duncan.LastName, Alias.LastName);
            Assert.AreEqual(Duncan.Affiliation, Alias.Affiliation);
            Assert.AreEqual(Duncan.Occupation, Alias.Occupation);

            // Compare data between Duncan and his Ghola
            Assert.AreEqual(Duncan.FirstName, Ghola.FirstName);
            Assert.AreEqual(Duncan.LastName, Ghola.LastName);
            Assert.AreEqual(Duncan.Affiliation, Ghola.Affiliation);
            Assert.AreNotEqual(Duncan.Occupation, Ghola.Occupation);

            // Switching an attribute in Duncan's alias switches it in Duncan but not the Ghola
            Alias.Occupation = "Ambassador to the Fremen";
            Assert.AreEqual(Duncan.Occupation, "Ambassador to the Fremen");
            Assert.AreEqual(Ghola.Occupation, "Ghola");
        }
    }
}
