using System;
using System.Collections.Generic;
using WordMerge.Models;
using Xunit;

namespace UnitTestWordMerge
{
    public class CoupleUnitTests
    {
        [Fact]
        public void ToString_ReturnsExpectedFormat()
        {
            var couple = new Couple<int,string>(1, "A");
            Assert.Equal("1 - A", couple.ToString());
        }

        [Fact]
        public void Equals_ReturnsTrueForEqualCouples()
        {
            var c1 = new Couple<int,string>(1, "A");
            var c2 = new Couple<int,string>(1, "A");
            Assert.True(c1.Equals(c2));
            Assert.True(c1 == c2);
        }

        [Fact]
        public void Equals_ReturnsFalseForDifferentCouples()
        {
            var c1 = new Couple<int,string>(1, "A");
            var c2 = new Couple<int,string>(2, "B");
            Assert.False(c1.Equals(c2));
            Assert.True(c1 != c2);
        }

        [Fact]
        public void GetHashCode_EqualCouples_HaveSameHashCode()
        {
            var c1 = new Couple<int,string>(1, "A");
            var c2 = new Couple<int,string>(1, "A");
            Assert.Equal(c1.GetHashCode(), c2.GetHashCode());
        }

        [Fact]
        public void Operators_CompareCorrectly()
        {
            var c1 = new Couple<int,string>(1, "A");
            var c2 = new Couple<int,string>(2, "B");
            Assert.True(c1 < c2);
            Assert.True(c2 > c1);
            Assert.True(c1 <= c2);
            Assert.True(c2 >= c1);
        }

        [Fact]
        public void ImplicitConversion_ToTuple_Works()
        {
            var couple = new Couple<int,string>(1, "A");
            Tuple<int, string> tuple = couple;
            Assert.Equal(1, tuple.Item1);
            Assert.Equal("A", tuple.Item2);
        }

        [Fact]
        public void ImplicitConversion_ToKeyValuePair_Works()
        {
            var couple = new Couple<int,string>(1, "A");
            KeyValuePair<int, string> kvp = couple;
            Assert.Equal(1, kvp.Key);
            Assert.Equal("A", kvp.Value);
        }

        [Fact]
        public void Handles_NullValues_Correctly()
        {
            var couple1 = new Couple<string, string>(null, "B");
            var couple2 = new Couple<string, string>("A", null);
            var couple3 = new Couple<string, string>(null, null);

            Assert.Equal(" - B", couple1.ToString());
            Assert.Equal("A - ", couple2.ToString());
            Assert.Equal(" - ", couple3.ToString());

            // Equality checks
            var couple4 = new Couple<string, string>(null, null);
            Assert.True(couple3.Equals(couple4));
            Assert.True(couple3 == couple4);
            Assert.Equal(couple3.GetHashCode(), couple4.GetHashCode());
        }
    }
}
