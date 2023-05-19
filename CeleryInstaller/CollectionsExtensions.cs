using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Unsafe = System.Runtime.CompilerServices.Unsafe;

namespace Masked.Sys.Extensions;

/// <summary>
/// Extensions for Collections
/// </summary>
public static class CollectionsExt {
    /// <summary>
    /// Iterates through an array fast, very fast.
    /// </summary>
    /// <typeparam name="T">Anything that is able to be in an Array.</typeparam>
    public static void FastIterator<T>(this T[] element, Func<T, int, NextStep> runInLoop) {
        ReadOnlySpan<T> spanOfElement = element;
        ref var searchSpace = ref MemoryMarshal.GetReference(spanOfElement);
        for (var i = 0; i < spanOfElement.Length; i++) {
            var obj = Unsafe.Add(ref searchSpace, i);
            var followingStep = runInLoop?.Invoke(obj, i);

            if (followingStep == NextStep.Break)
                break;
        }
    }

    /// <summary>
    /// Iterates through a List fast, very fast.
    /// </summary>
    /// <typeparam name="T">Anything that is able to be in a List.</typeparam>
    /// <remarks>The <see cref="IEnumerable{T}"/> can NOT be modified! Else this operation provide incorrect results and will be dangerous to use.</remarks>
    public static void FastIterator<T>(this IEnumerable<T> element, Func<T, int, NextStep> runInLoop) {
        Span<T> spanOfElement = CollectionsMarshal.AsSpan(element.ToList());
        ref var searchSpace = ref MemoryMarshal.GetReference(spanOfElement);
        for (var i = 0; i < spanOfElement.Length; i++) {
            var obj = Unsafe.Add(ref searchSpace, i);
            var followingStep = runInLoop?.Invoke(obj, i);

            if (followingStep == NextStep.Break)
                break;
        }
    }
}

/// <summary>
/// Enumeration used to communicate to the <see cref="CollectionsExt.FastIterator{T}(IEnumerable{T}, Func{T, int, NextStep})"/> and <see cref="CollectionsExt.FastIterator{T}(T[], Func{T, int, NextStep})"/> methods if the loop should end, or continue
/// </summary>
public enum NextStep {
    /// <summary>
    /// Should the iterator break out of the loop after this statement
    /// </summary>
    Break,

    /// <summary>
    /// Should the iterator continue looping.
    /// </summary>
    Continue
}