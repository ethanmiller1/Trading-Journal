# FORMAT

## Format column A

Use conditional formatting make the color of column A indicate its market posture.

```excel
=GetPosture(GetStrategy($A7,$D7))="Neutral"
=GetPosture(GetStrategy($A7,$D7))="Bullish"
=GetPosture(GetStrategy($A7,$D7))="Bearish"
```

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/postureFormat.gif)

## Display date as month only

`Right-click` > `Format Cells` > `Custom`, and then paste in:

```excel
[$-en-US]mmm;
```

## Shade every other row

Create a new rule in conditional formatting, select `Use a formula to determine which cells to format`, and then:

```excel
=MOD(ROW(),2)=1
```

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/colorShade.png)

I use rgb(234, 238, 225)

## Add cell padding

`Home` > `Alignment` > `Increase Indent`

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/cellPadding.gif)

## Add column dividers

`Border` > `Draw Borders`

![](https://i.ibb.co/gZ4NMgV/image.png)

I use rgb(211, 221, 196) and the thickest line style

## Display 0% as blank

`Right-click` > `Format Cells` > `Custom`, and then paste in:

```excel
0.00%;(0.00%);""
```

![](https://i.ibb.co/dDPbQb1/image.png)

# Dashbord

## Double Bottom Border

![](https://i.ibb.co/bgrhFCF/image.png)

I use `#A6A6A6` and the double bottom border. The font is all caps `Verdana` with a font size of `9` and `Red, Accent 2`.

![](https://i.ibb.co/7GpFKfg/image.png)

## Horizontal Bar Chart

![](https://i.ibb.co/nj0MLf1/image.png)

![](https://i.ibb.co/2gMgxYH/image.png)

![](https://i.ibb.co/1Kt9WjQ/image.png)

![](https://i.ibb.co/rkNq4Dq/image.png)

### Format Gridlines

![](https://i.ibb.co/2Ktw8Y0/image.png)

![](https://i.ibb.co/MM9MGsK/image.png)

### Remove Chart Border

![](https://i.ibb.co/8zG61x1/image.png)

### Format Axis

![](https://i.ibb.co/mDcw78q/image.png)

### Display negative and positive numbers in same direction

This can be achieved by converting the loss column to positive values. Use the following format to indicate they are negative numbers.

```excel
_($* (#,##0)" ";_($* #,##0_)" ";"";_(@_)" "
```

![](https://i.ibb.co/tM1Vn7p/image.png)

### Choose specific label intervals

![](https://i.ibb.co/HTCcDDM/image.png)

## Thermometer bar

Select 4 cells individually.

![](https://i.ibb.co/qJQrz33/image.png)

![](https://i.ibb.co/H7ffgqh/Untitled.png)

### Display as percentage

Change Chart Type to Stacked Column, select the small data series, and change to Secondary Axis.

![](https://i.ibb.co/68kcC9s/image.png)

### Change color

![](https://i.ibb.co/NrtsWrL/image.png)

### Add Tick Marks

![](https://i.ibb.co/MMfDYYg/image.png)

### Resize

![](https://i.ibb.co/6tpZvtF/image.png)
