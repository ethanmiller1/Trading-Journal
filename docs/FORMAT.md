# FORMAT

## Format column A

Use conditional formatting make the color of column A indicate its market posture.

``` excel
=getPosture(getStrategy($A7,$D7))="Neutral"
=getPosture(getStrategy($A7,$D7))="Bullish"
=getPosture(getStrategy($A7,$D7))="Bearish"
```

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/postureFormat.gif)

## Display date as month only

`Right-click` > `Format Cells` > `Custom`, and then paste in:

``` excel
[$-en-US]mmm;
```

## Shade every other row

Create a new rule in conditional formatting, select `Use a formula to determine which cells to format`, and then:

``` excel
=MOD(ROW(),2)=1
```

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/colorShade.png)

I use rgb(234, 238, 225) 

## Add cell padding

`Home` > `Alignment` > `Increase Indent`

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/cellPadding.png)