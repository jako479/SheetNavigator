using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SheetNavigator.Tests
{
    /// <summary>
    /// The width and height rules at their limits: the smallest accepted value, one below it,
    /// the ceiling, and one past it.
    /// </summary>
    [TestClass]
    public class PaneSizeRulesTests
    {
        /// <summary>A value no rule produces on its own, so a returned fallback is unmistakable.</summary>
        private const int Fallback = 99;

        [TestMethod]
        public void ClampWidth_SmallestPositive_IsKept()
        {
            Assert.AreEqual(1, PaneSizeRules.ClampWidth(1, Fallback));
        }

        [TestMethod]
        public void ClampWidth_Zero_IsTheFallback()
        {
            Assert.AreEqual(Fallback, PaneSizeRules.ClampWidth(0, Fallback));
        }

        [TestMethod]
        public void ClampWidth_AtTheCeiling_IsKept()
        {
            Assert.AreEqual(400, PaneSizeRules.ClampWidth(400, Fallback));
        }

        [TestMethod]
        public void ClampWidth_OnePastTheCeiling_IsTheCeiling()
        {
            Assert.AreEqual(400, PaneSizeRules.ClampWidth(401, Fallback));
        }

        [TestMethod]
        public void ClampHeight_SmallestPositive_IsKept()
        {
            Assert.AreEqual(1, PaneSizeRules.ClampHeight(1, Fallback));
        }

        [TestMethod]
        public void ClampHeight_Zero_IsTheFallback()
        {
            Assert.AreEqual(Fallback, PaneSizeRules.ClampHeight(0, Fallback));
        }

        [TestMethod]
        public void ClampHeight_AtTheCeiling_IsKept()
        {
            Assert.AreEqual(1200, PaneSizeRules.ClampHeight(1200, Fallback));
        }

        [TestMethod]
        public void ClampHeight_OnePastTheCeiling_IsTheCeiling()
        {
            Assert.AreEqual(1200, PaneSizeRules.ClampHeight(1201, Fallback));
        }
    }
}
