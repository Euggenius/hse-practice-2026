using System;
using System.Collections.Generic;
using System.Linq;

namespace GostEspdAddIn.Services
{
    internal static class FloatEnumerableExtensions
    {
        public static IEnumerable<float>
            DistinctWithTolerance(
                this IEnumerable<float> source,
                float tolerance)
        {
            var result =
                new List<float>();

            foreach (float value in source)
            {
                bool exists =
                    result.Any(
                        x =>
                            Math.Abs(
                                x - value) <=
                            tolerance);

                if (!exists)
                    result.Add(value);
            }

            return result;
        }
    }
}