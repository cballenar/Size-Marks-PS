/**
 * Size Marks 1.3
 *
 * Copyright (c) 2014 Roman Shamin https://github.com/romashamin
 * and licenced under the MIT licence. All rights not explicitly
 * granted in the MIT license are reserved. See the included
 * LICENSE file for more details.
 *
 * https://github.com/romashamin
 * https://twitter.com/romanshamin
 *
 * Converts rectangular selection to labeled measurement mark.
 * Landscape selection → horizontal mark. Portrait or square
 * selection → vertical mark.
 */


var doc = null,
    docIsExist = false,
    selBounds = null,
    selIsExist = false;

var store = {
    activeLayer: null,
    rulerUnits: app.preferences.rulerUnits,
    typeUnits: app.preferences.typeUnits,
    font: null
};


try {
    doc = app.activeDocument;
    docIsExist = true;
} catch (e) {
    alert('Size Marks Script: no document\n' +
          'Use File → New... to create one');
}


if (docIsExist) {
    try {
        selBounds = doc.selection.bounds;
        selIsExist = true;
    } catch (e) {
        alert('Size Marks Script: no selection\n' +
              'Use Rectangular Marquee Tool (M) to create one');
    }
}


// get realValues.unit, realValues.width, and realValues.height
var realValues = getRealValues(app.preferences.rulerUnits, selBounds);

// normalize document by setting working units to pixels and points
app.preferences.rulerUnits = Units.PIXELS;
app.preferences.typeUnits = TypeUnits.POINTS;

// reset selection using working units
selBounds = doc.selection.bounds;


if (docIsExist && selIsExist) {
    doc.suspendHistory("Add Size Mark", "makeSizeMark()");
}


function makeSizeMark() {
  try {
    var baseRes = 72,
        layerOpacity = 65,
        docRes = doc.resolution,
        scaleRatio = docRes / baseRes,
        scale = setScaleF(scaleRatio),
        charThinSpace = '\u200A'; /* Thin space: \u2009, hair space: \u200A */

    // Values relative to resolution
    var lineWidth = 1 * round(scaleRatio, 0),
        halfMark  = 3 * round(scaleRatio, 0),
        txtMargin = 6 * round(scaleRatio, 0);

    var selX1 = selBounds[0].value,
        selX2 = selBounds[2].value,
        selY1 = selBounds[1].value,
        selY2 = selBounds[3].value;

    var selWidth = selX2 - selX1,
        selHeight = selY2 - selY1;

    var val = 0,
        realVal = 0,
        txtLayerPos = [0, 0],
        layerNamePrefix = 'MSRMNT',
        txtJ11n = Justification.LEFT;

    store.activeLayer = doc.activeLayer;
    doc.selection.deselect();
    var markLayer = doc.artLayers.add();

    setPenToolSize(lineWidth);

    if (selWidth > selHeight) {
        // Adjust points based on line width
        // Useful for resolution other than 72ppi
        adjSelX1 = selX1 + lineWidth/2;
        adjSelX2 = selX2 - lineWidth/2;

        // Draw Main Line
        drawLine([adjSelX1, selY1], [adjSelX2, selY1]);

        // Draw Edge Marks
        drawLine([adjSelX1, selY1 - halfMark], [adjSelX1, selY1 + halfMark]);
        drawLine([adjSelX2, selY1 - halfMark], [adjSelX2, selY1 + halfMark]);

        // Set some values for text layer
        layerNamePrefix = 'W';
        val = selWidth;
        realVal = realValues.width;
        txtLayerPos = [selX1 + val / 2, selY1 - txtMargin];
        txtJ11n = Justification.CENTER;

    } else {
        // Adjust points based on line width
        // Useful for resolution greater than 72ppi
        adjSelY1 = selY1 + lineWidth/2;
        adjSelY2 = selY2 - lineWidth/2;

        // Draw Main Line
        drawLine([selX1, adjSelY1], [selX1, adjSelY2]);

        // Draw Edge Marks
        drawLine([selX1 - halfMark, adjSelY1], [selX1 + halfMark, adjSelY1]);
        drawLine([selX1 - halfMark, adjSelY2], [selX1 + halfMark, adjSelY2]);

        // Set some values for text layer
        layerNamePrefix = 'H';
        val = selHeight;
        realVal = realValues.height;
        txtLayerPos = [selX1 + txtMargin, selY1 + val / 2 + 4];
        txtJ11n = Justification.LEFT;
    }

    markLayer.opacity = 85;
    markLayer.move(store.activeLayer, ElementPlacement.PLACEBEFORE);

    // Draw label
    disableArtboardAutoNest();

    var txtLayer = makeTextLayer(),
        txtLayerItem = txtLayer.textItem;

    store.font = txtLayerItem.font;

    txtLayerItem.size = 12;
    txtLayerItem.font = 'ArialMT';
    txtLayerItem.autoKerning = AutoKernType.OPTICAL;

    txtLayer.translate(txtLayerPos[0], txtLayerPos[1]);

    txtLayerItem.justification = txtJ11n;
    txtLayerItem.color = app.foregroundColor;

    var label = formatValueWithUnits(realVal, realValues.unit, charThinSpace);
    txtLayerItem.contents = label;

    // Finish
    txtLayer.rasterize(RasterizeType.TEXTCONTENTS);
    txtLayer.move(markLayer, ElementPlacement.PLACEBEFORE);

    var finalLayer = txtLayer.merge();
    finalLayer.name = layerNamePrefix + ' ' + label;
    finalLayer.opacity = layerOpacity;

    app.preferences.rulerUnits = store.rulerUnits;
    app.preferences.typeUnits = store.typeUnits;

    enableArtboardAutoNest();

    pickTool('marqueeRectTool');

    // HELPERS

    function makePoint(pnt) {

        for (var i = 0; i < pnt.length; i++) {
            pnt[i] = scale(pnt[i]);
        }

        var point = new PathPointInfo();

        point.anchor = pnt;
        point.leftDirection = pnt;
        point.rightDirection = pnt;
        point.kind = PointKind.CORNERPOINT;

        return point;
    }

    function setScaleF(ratio) {
        return function (value) {
            return value / ratio;
        }
    }


    function formatValueWithUnits(v, u, space) {
        space = space || '';
        return '' + v + space + u;
    }


    function drawLine(start, stop) {

        var startPoint = makePoint(start),
            stopPoint = makePoint(stop);

        var spi = new SubPathInfo();
        spi.closed = false;
        spi.operation = ShapeOperation.SHAPEXOR;
        spi.entireSubPath = [startPoint, stopPoint];

        var uniqueName = 'Line ' + Date.now();
        var line = doc.pathItems.add(uniqueName, [spi]);
        line.strokePath(ToolType.PENCIL);
        line.remove();
    }


    function pickTool(toolName) {
        var idslct = charIDToTypeID('slct');
        var desc4 = new ActionDescriptor();
        var idnull = charIDToTypeID('null');
        var ref2 = new ActionReference();
        var idmarqueeRectTool = stringIDToTypeID(toolName);
        ref2.putClass(idmarqueeRectTool);
        desc4.putReference(idnull, ref2);
        executeAction(idslct, desc4, DialogModes.NO);
    }


    function makeTextLayer() {
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putClass(app.charIDToTypeID('TxLr'));
        desc.putReference(app.charIDToTypeID('null'), ref);
        var desc2 = new ActionDescriptor();
        desc2.putString(app.charIDToTypeID('Txt '), "text");
        var list2 = new ActionList();
        desc2.putList(app.charIDToTypeID('Txtt'), list2);
        desc.putObject(app.charIDToTypeID('Usng'), app.charIDToTypeID('TxLr'), desc2);
        executeAction(app.charIDToTypeID('Mk  '), desc, DialogModes.NO);
        return doc.activeLayer
    }


    function disableArtboardAutoNest() {
        var ideditArtboardEvent = stringIDToTypeID( "editArtboardEvent" );
        var desc3 = new ActionDescriptor();
        var idnull = charIDToTypeID( "null" );
            var ref2 = new ActionReference();
            var idLyr = charIDToTypeID( "Lyr " );
            var idOrdn = charIDToTypeID( "Ordn" );
            var idTrgt = charIDToTypeID( "Trgt" );
            ref2.putEnumerated( idLyr, idOrdn, idTrgt );
        desc3.putReference( idnull, ref2 );
        var idautoNestEnabled = stringIDToTypeID( "autoNestEnabled" );
        desc3.putBoolean( idautoNestEnabled, false );
        executeAction( ideditArtboardEvent, desc3, DialogModes.NO );
    }


    function enableArtboardAutoNest() {
        var ideditArtboardEvent = stringIDToTypeID( "editArtboardEvent" );
        var desc3 = new ActionDescriptor();
        var idnull = charIDToTypeID( "null" );
            var ref2 = new ActionReference();
            var idLyr = charIDToTypeID( "Lyr " );
            var idOrdn = charIDToTypeID( "Ordn" );
            var idTrgt = charIDToTypeID( "Trgt" );
            ref2.putEnumerated( idLyr, idOrdn, idTrgt );
        desc3.putReference( idnull, ref2 );
        var idautoNestEnabled = stringIDToTypeID( "autoNestEnabled" );
        desc3.putBoolean( idautoNestEnabled, true );
        executeAction( ideditArtboardEvent, desc3, DialogModes.NO );
    }


    /**
     * Source: https://forums.adobe.com/thread/962285?start=0&tstart=0
     * Comment for Feb 16, 2012 7:18 AM
     */
    function setPenToolSize(dblSize) {
        var idslct = charIDToTypeID('slct');
        var desc3 = new ActionDescriptor();
        var idnull = charIDToTypeID('null');
        var ref2 = new ActionReference();
        var idPcTl = charIDToTypeID('PcTl');
        ref2.putClass(idPcTl);
        desc3.putReference(idnull, ref2);
        executeAction(idslct, desc3, DialogModes.NO);

        var idsetd = charIDToTypeID('setd');
        var desc2 = new ActionDescriptor();
        var ref1 = new ActionReference();
        var idBrsh = charIDToTypeID('Brsh');
        var idOrdn = charIDToTypeID('Ordn');
        var idTrgt = charIDToTypeID('Trgt');
        ref1.putEnumerated(idBrsh, idOrdn, idTrgt);
        desc2.putReference(idnull, ref1);
        var idT = charIDToTypeID('T   ');
        var idmasterDiameter = stringIDToTypeID('masterDiameter');
        var idPxl = charIDToTypeID('#Pxl');
        desc3.putUnitDouble(idmasterDiameter, idPxl, dblSize);
        desc2.putObject(idT, idBrsh, desc3);
        executeAction(idsetd, desc2, DialogModes.NO);
    }
  } catch (e) {
    alert(e.line + '\n' + e)
  }
}

function getRealValues(rulerUnits, selectionBouds) {
    var realValues = {};

    // identify unit
    if ( rulerUnits == Units.PIXELS ) {
        realValues.unit = 'px';
        realValues.decimals = 0;
    } else if ( rulerUnits == Units.INCHES ) {
        realValues.unit = 'in';
        realValues.decimals = 3;
    } else if ( rulerUnits == Units.CM ) {
        realValues.unit = 'cm';
        realValues.decimals = 1;
    } else if ( rulerUnits == Units.MM ) {
        realValues.unit = 'mm';
        realValues.decimals = 0;
    } else if ( rulerUnits == Units.POINTS ) {
        realValues.unit = 'pt';
        realValues.decimals = 1;
    } else if ( rulerUnits == Units.PICAS ) {
        realValues.unit = 'pc';
        realValues.decimals = 2;
    }

    // define selection coordinates
    var realSelX1 = selectionBouds[0].value,
        realSelX2 = selectionBouds[2].value,
        realSelY1 = selectionBouds[1].value,
        realSelY2 = selectionBouds[3].value;

    // get width and height in real units
    realValues.width = round(realSelX2 - realSelX1, realValues.decimals);
    realValues.height = round(realSelY2 - realSelY1, realValues.decimals);

    return realValues;
}

function round(value, decimals) {
    return Number(Math.round(value+'e'+decimals)+'e-'+decimals);
}
