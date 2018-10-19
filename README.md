# VBA-StringBuilder

A `StringBuilder` class based on the work of @Blackhawk, published under CC-BY-SA on [Code Review Stack Exchange](https://codereview.stackexchange.com/q/67596/23788).

Concatenating strings in a loop is a very inefficient process. Using a `StringBuilder` for concatenating a large number of strings makes the code much more efficient and noticeably enhances performance.

- 1K strings; both concat & builder are pretty much instant
- 10K strings; both concat & builder complete in ~0.05 seconds
- 20K strings; very similar performance, ~0.1 seconds for both
- 50K strings; concat ~0.65 seconds, builder ~0.2 seconds
- 100K strings; concat ~2.7 seconds, builder ~0.4 seconds
- 1M strings; concat ????\*, builder ~4.9 seconds

<sub>\*Gave up, killed the process after 10 minutes.</sub>

## Usage

With the `StringBuilder` class added to your VBA project, use the `New` keyword to create a new instance of the class:

```vb
Dim sb As StringBuilder
Set sb = New StringBuilder
```

This initializes the builder with a default initial capacity.
To specify an initial capacity and/or initial content, or to create an instance of the class from a VBA project that's referencing the VBA project the `StringBuilder` class is loaded from (e.g. an add-in), use the `Create` factory method off the *default instance* instead:

```vb
Dim sb As StringBuilder
Set sb = StringBuilder.Create("test", 32)
```

Use the `Append` method to, well, *append* a string to the builder; use the `ToString` method to retrieve the string:

```vb
Dim sb As StringBuilder
Set sb = StringBuilder.Create

Dim i As Long
For i = 1 To 100000
    sb.Append "Test"
Next
Debug.Print sb.ToString
```

Instance members cannot be invoked from the *default instance*.
