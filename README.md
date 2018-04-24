# ExcelProgress

A simple progress indicator that works the same way as Access SysCmd.

![Imgur](https://i.imgur.com/D5XMmB2.png)
![Imgur](https://i.imgur.com/hj5Tsi2.png)
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

Works VERY well with [Reckon](https://github.com/groovyMysterioso/Reckon)
```
ProgressBar xUpdateMeter, EstimateTick(indexNumber,whatever.Count), indexNumber
```

## Authors

* **James Pritts** - *Initial work* - [GroovyMysterioso](https://github.com/GroovyMysterioso)

