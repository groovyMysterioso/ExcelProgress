# ExcelProgress

A simple progress indicator that works the same way as Access SysCmd.

## Getting Started

Simply import into your Workbook or Personal Workbook in /XLSTART!

Initialize the progress bar as in Access

```
ProgressBar xlInitMeter, "Initializing", whatEver.Count
```

Then update

```
ProgressBar xlUpdateMeter, "Completed", indexNumber
```

Finally destroy the progress bar

```
ProgressBar xlRemoveMeter
```


## Authors

* **James Pritts** - *Initial work* - [GroovyMysterioso](https://github.com/GroovyMysterioso)

