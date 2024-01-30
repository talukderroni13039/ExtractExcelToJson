```

BenchmarkDotNet v0.13.12, Windows 11 (10.0.22621.2861/22H2/2022Update/SunValley2)
11th Gen Intel Core i7-1165G7 2.80GHz, 1 CPU, 8 logical and 4 physical cores
  [Host]     : .NET Framework 4.8.1 (4.8.9181.0), X86 LegacyJIT [AttachedDebugger]
  DefaultJob : .NET Framework 4.8.1 (4.8.9181.0), X86 LegacyJIT


```
| Method                   | Mean        | Error     | StdDev    | Median      |
|------------------------- |------------:|----------:|----------:|------------:|
| ExtractDatafromFastExcel |    47.19 ms |  0.931 ms |  1.243 ms |    47.07 ms |
| ExtractDatafromEPPlus    |    28.93 ms |  0.603 ms |  1.681 ms |    28.69 ms |
| ExportDataFromSpire      | 1,022.56 ms | 20.110 ms | 41.976 ms | 1,006.81 ms |
