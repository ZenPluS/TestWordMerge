using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using WordMerge.Core;
using WordMerge.Extensions;
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

        [Fact]
        public void IStructuralEquatable_Equals_UsesCustomComparer()
        {
            var c1 = new Couple<string, string>("a", "b");
            var c2 = new Couple<string, string>("A", "B");
            var comparer = StringComparer.OrdinalIgnoreCase;

            IStructuralEquatable structural = c1;
            Assert.True(structural.Equals(c2, comparer));
        }

        [Fact]
        public void IStructuralEquatable_GetHashCode_UsesCustomComparer()
        {
            var c1 = new Couple<string, string>("a", "b");
            var c2 = new Couple<string, string>("A", "B");
            var comparer = StringComparer.OrdinalIgnoreCase;

            IStructuralEquatable structural = c1;
            Assert.Equal(structural.GetHashCode(comparer), c2.GetHashCode(comparer));
        }

        [Fact]
        public void IStructuralComparable_CompareTo_UsesCustomComparer()
        {
            var c1 = new Couple<string, string>("a", "b");
            var c2 = new Couple<string, string>("A", "c");
            var comparer = StringComparer.OrdinalIgnoreCase;

            IStructuralComparable structural = c1;
            Assert.True(structural.CompareTo(c2, comparer) < 0);
        }

        [Fact]
        public void ITuple_IndexerAndLength_WorkCorrectly()
        {
            var c = new Couple<int, string>(42, "foo");
            ITuple tuple = c;

            Assert.Equal(2, tuple.Length);
            Assert.Equal(42, tuple[0]);
            Assert.Equal("foo", tuple[1]);
            Assert.Throws<IndexOutOfRangeException>(() => tuple[2]);
        }

         [Fact]
        public void CompareTo_ReturnsOne_WhenOtherIsNull()
        {
            var c = new Couple<int, string>(5, "X");
            Assert.True(c.CompareTo(null) > 0);
        }

        [Fact]
        public void ComparisonOperators_HandleNulls()
        {
            Couple<int,string> a = null;
            var b = new Couple<int,string>(1, "B");

            Assert.True(a == null);
            Assert.True(b != null);
            Assert.True(a < b);
            Assert.False(b < a);
            Assert.True(b > a);
            Assert.True(a <= b);
            Assert.True(b >= a);
        }

        [Fact]
        public void CompareTo_UsesLeftThenRight()
        {
            var c1 = new Couple<int,string>(1, "B");
            var c2 = new Couple<int,string>(2, "A");
            var c3 = new Couple<int,string>(1, "C");

            Assert.True(c1.CompareTo(c2) < 0); // Left decides
            Assert.True(c1.CompareTo(c3) < 0); // Left equal, Right decides
        }

        [Fact]
        public void IsEqualTo_ReturnsTrue_ForSameValues()
        {
            var c1 = new Couple<int,string>(1,"A");
            var c2 = new Couple<int,string>(1,"A");
            Assert.True(c1.IsEqualTo(c2));
        }

        [Fact]
        public void IsEqualTo_Throws_WhenOtherNull()
        {
            var c1 = new Couple<int,string>(1,"A");
            Assert.Throws<ArgumentNullException>(() => c1.IsEqualTo(null));
        }

        [Fact]
        public void LeftItems_And_RightItems_Work()
        {
            var list = new List<ICouple<int,string>>
            {
                new Couple<int,string>(1,"A"),
                new Couple<int,string>(2,"B")
            };

            Assert.Equal(new[]{1,2}, list.LeftItems());
            Assert.Equal(new[]{"A","B"}, list.RightItems());
        }
    }
}
