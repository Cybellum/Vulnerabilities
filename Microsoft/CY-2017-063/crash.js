(function() {
    "use strict";
    // The initialize function is run each time the page is loaded.
    Office.initialize = function(reason) {
        $(document).ready(function() {

            // Use this to check whether the API is supported in the app client.
            if (Office.context.requirements.isSetSupported('ExcelApi', 1.6)) {
                $('#crashButton').click(crashCreate);
                $('#supportedVersion').html('Ready');
            } else {
                // Just letting you know that this code will not work with your app version.
                $('#supportedVersion').html('This code requires latest version of the application');
            }
        });
    };

    function crashCreate() {
        return Excel.run(function(context) {
            var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
            context.load(range);
            cyVar0 = range;
            return context.sync()
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar0 = cyVar0.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar1 = tempVar0;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1 = cyVar0.getRow(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar2 = tempVar1;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar2 = cyVar0.getBoundingRect('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar3 = tempVar2;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar3 = cyVar0.track(false, 1, 34343333, false, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar4 = cyVar0.set(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar5 = cyVar0.getLastColumn(false, true, null, 'asf', null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar4 = tempVar5;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar6 = cyVar0.getLastCell(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar5 = tempVar6;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar7 = cyVar0.getUsedRange(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar6 = tempVar7;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar8 = cyVar0.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar7 = tempVar8;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar9 = cyVar0.untrack(1, -1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar10 = cyVar0._ValidateArraySize(-1, 34343333, true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar11 = cyVar0.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar12 = cyVar0.getRowsAbove(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar13 = cyVar0.untrack(undefined, 3.5, 3.5, true, 1, true, true, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar14 = cyVar0.getColumn(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar15 = cyVar0._ensureInteger(false);
                    } catch (err) {}
                    try {
                        tempVar16 = cyVar0._ensureInteger(false, 34343333, null, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar17 = cyVar0.unmerge(3.5, 0, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar18 = cyVar0.getLastRow(34343333, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar19 = cyVar0.getCell(1, true, false, 'asf', undefined, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar20 = cyVar0.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar21 = cyVar0.getResizedRange(5, 13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar22 = cyVar0.toJSON('asf', 1, true, 0, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar8 = tempVar22;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar23 = cyVar0.getColumnsAfter(34343333, 1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar24 = cyVar0._handleResult(undefined, -1, 'asf', -1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar25 = cyVar0.getRow(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar9 = tempVar25;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar26 = cyVar0.calculate(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar27 = cyVar0.merge(false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar28 = cyVar0._handleIdResult(false, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar29 = cyVar0._ValidateArraySize(3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar30 = cyVar0._getAdjacentRange(null, null, undefined, 34343333);
                    } catch (err) {}
                    try {
                        tempVar31 = cyVar0._getAdjacentRange('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar32 = cyVar0.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar10 = tempVar32;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar33 = cyVar0.getEntireColumn(3.5, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar11 = tempVar33;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar34 = cyVar0.select(null, 34343333, 1, 0, 0, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar35 = cyVar0.getLastColumn(false, 3.5, true, 'asf', 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar12 = tempVar35;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar36 = cyVar0.getRowsBelow(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar13 = tempVar36;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar37 = cyVar0.set('asf', 1, undefined, 34343333, 1);
                    } catch (err) {}
                    try {
                        tempVar38 = cyVar0.set('asf', undefined, false, undefined, 0, false, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar39 = cyVar0.getOffsetRange(1, -1, 0, 3.5, 1, 34343333, undefined, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar14 = tempVar39;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar40 = cyVar0.getVisibleView(false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar15 = tempVar40;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar41 = cyVar0.getUsedRange(0, 34343333, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar16 = tempVar41;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar42 = cyVar0._KeepReference(0, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar43 = cyVar0.getLastCell('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar17 = tempVar43;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar44 = cyVar0._recursivelySet(undefined, -1, undefined, true, true, 3.5, 3.5);
                    } catch (err) {}
                    try {
                        tempVar45 = cyVar0._recursivelySet(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar46 = cyVar0.getColumnsBefore(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar18 = tempVar46;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar47 = cyVar0.track(1, 3.5, undefined, null, false, true, -1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar48 = cyVar0.getEntireRow(3.5, 0, 3.5, true, false, 3.5, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar19 = tempVar48;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar49 = cyVar19.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar20 = tempVar49;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar50 = cyVar19._getAdjacentRange(true, 3.5, true, undefined, 'asf', true, 1, null);
                    } catch (err) {}
                    try {
                        tempVar51 = cyVar19._getAdjacentRange(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar52 = cyVar19._handleResult(false, 'asf', 'asf', 3.5, true, null, 0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar53 = cyVar19.toJSON(3.5, null, 3.5, undefined, -1, 1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar21 = tempVar53;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar54 = cyVar19.getColumnsBefore(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar22 = tempVar54;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar55 = cyVar19.getEntireRow(0, null, false, false, 0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar23 = tempVar55;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar56 = cyVar23.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar24 = tempVar56;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar57 = cyVar23.getColumnsBefore(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar25 = tempVar57;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar58 = cyVar23.getRow(1, 'asf', 34343333, 1, 3.5, 'asf', 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar26 = tempVar58;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar59 = cyVar23.unmerge(undefined, false, true, 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar60 = cyVar23.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar27 = tempVar60;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar61 = cyVar23.load(0, null, 0, null, 'asf');
                    } catch (err) {}
                    try {
                        tempVar62 = cyVar23.load(true, 34343333, 3.5, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar63 = cyVar23.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar28 = tempVar63;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar64 = cyVar23.getVisibleView(-1, 'asf', true, 34343333, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar29 = tempVar64;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar65 = cyVar23.untrack(undefined, 3.5, 1, -1, true, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar66 = cyVar23.getEntireRow('asf', 'asf', false, undefined, 0, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar30 = tempVar66;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar67 = cyVar23._getAdjacentRange(1, true, undefined, 3.5);
                    } catch (err) {}
                    try {
                        tempVar68 = cyVar23._getAdjacentRange(true, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar69 = cyVar23.merge(-1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar70 = cyVar22.getLastColumn(34343333, -1, null, 1, undefined, undefined, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar31 = tempVar70;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar71 = cyVar22.getVisibleView('asf', true, 'asf', true, 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar32 = tempVar71;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar72 = cyVar22.unmerge(0, 3.5, 3.5, 'asf', null, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar73 = cyVar22.getColumnsAfter(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar33 = tempVar73;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar74 = cyVar22.getCell(8, 2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar34 = tempVar74;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar75 = cyVar22.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar35 = tempVar75;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar76 = cyVar20.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar36 = tempVar76;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar77 = cyVar20.getVisibleView(undefined, 1, -1, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar37 = tempVar77;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar78 = cyVar20._getAdjacentRange(3.5, 34343333, 3.5, undefined, 0, true, true);
                    } catch (err) {}
                    try {
                        tempVar79 = cyVar20._getAdjacentRange(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar80 = cyVar20.load(0, true);
                    } catch (err) {}
                    try {
                        tempVar81 = cyVar20.load(null, 'asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar82 = cyVar20.getRow(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar38 = tempVar82;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar83 = cyVar20.getColumn(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar39 = tempVar83;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar84 = cyVar20.getOffsetRange(-1, 'asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar40 = tempVar84;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar85 = cyVar20._ValidateArraySize(-1, null, 34343333, true, 0, 1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar86 = cyVar20._recursivelySet(34343333, 1, 0);
                    } catch (err) {}
                    try {
                        tempVar87 = cyVar20._recursivelySet(0, -1, null, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar88 = cyVar20.getRowsBelow(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar41 = tempVar88;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar89 = cyVar20.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar42 = tempVar89;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar90 = cyVar20.select(34343333, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar91 = cyVar20.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar43 = tempVar91;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar92 = cyVar20.getLastCell(1, 'asf', 'asf', -1, 3.5, 1, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar44 = tempVar92;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar93 = cyVar20.getIntersection(34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar45 = tempVar93;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar94 = cyVar20.getLastColumn(34343333, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar46 = tempVar94;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar95 = cyVar20.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar47 = tempVar95;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar96 = cyVar20._KeepReference(34343333, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar97 = cyVar20.getEntireRow(3.5, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar48 = tempVar97;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar98 = cyVar20.getCell(3, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar49 = tempVar98;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar99 = cyVar20.set(0, -1, null, 0, 'asf', null, false, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar100 = cyVar20.getColumnsAfter(8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar50 = tempVar100;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar101 = cyVar20.merge(false, 'asf', true, 0, -1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar102 = cyVar20.calculate(undefined, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar103 = cyVar20.untrack(false, false, undefined, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar104 = cyVar20.getColumnsBefore(8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar51 = tempVar104;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar105 = cyVar20.getRowsAbove(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar52 = tempVar105;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar106 = cyVar20.getUsedRange(null, false, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar53 = tempVar106;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar107 = cyVar20.track(-1, true, 34343333, false, 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar108 = cyVar20.getLastRow(1, 34343333, 1, null, null, false, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar54 = tempVar108;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar109 = cyVar20._handleResult(1, 3.5, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar110 = cyVar20.toJSON(-1, 'asf', -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar55 = tempVar110;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar111 = cyVar20.getResizedRange(13, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar56 = tempVar111;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar112 = cyVar20.unmerge('asf', true, undefined, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar113 = cyVar20._ensureInteger(3.5, 1, false, false);
                    } catch (err) {}
                    try {
                        tempVar114 = cyVar20._ensureInteger(false, 34343333, 34343333, 0, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar115 = cyVar20._handleIdResult(3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar116 = cyVar18.merge(0, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar117 = cyVar18.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar57 = tempVar117;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar118 = cyVar18.getEntireRow(undefined, true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar58 = tempVar118;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar119 = cyVar18.getRow(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar59 = tempVar119;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar120 = cyVar18.getLastCell(1, -1, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar60 = tempVar120;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar121 = cyVar18.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar61 = tempVar121;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar122 = cyVar18.getColumn(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar62 = tempVar122;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar123 = cyVar18.select(null, 3.5, 34343333, true, 'asf', 'asf', -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar124 = cyVar18.load(undefined, false, 1, 'asf', 'asf', null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar125 = cyVar18._recursivelySet(1, 'asf', null, true);
                    } catch (err) {}
                    try {
                        tempVar126 = cyVar18._recursivelySet(0, undefined, undefined, 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar127 = cyVar18.getOffsetRange(9, 6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar63 = tempVar127;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar128 = cyVar18.calculate(true, false, undefined, -1, undefined, undefined, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar129 = cyVar18.set(false, undefined, undefined, 34343333, undefined, 34343333, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar130 = cyVar18.getCell(4, 8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar64 = tempVar130;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar131 = cyVar18.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar65 = tempVar131;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar132 = cyVar18.getVisibleView(true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar66 = tempVar132;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar133 = cyVar18.getColumnsBefore(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar67 = tempVar133;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar134 = cyVar18._KeepReference(34343333, 'asf', -1, undefined, false, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar135 = cyVar18.getLastRow(3.5, 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar68 = tempVar135;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar136 = cyVar18._handleResult(3.5, 3.5, 3.5, -1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar137 = cyVar18.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar69 = tempVar137;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar138 = cyVar18.untrack(0, 3.5, 1, 'asf', undefined, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar139 = cyVar18.toJSON('asf', null, 34343333, 34343333, 1, 'asf', 0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar70 = tempVar139;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar140 = cyVar18._ensureInteger(true, 3.5, true, 1);
                    } catch (err) {}
                    try {
                        tempVar141 = cyVar18._ensureInteger('asf', -1, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar142 = cyVar18._getAdjacentRange(34343333, 'asf', 1, 'asf', -1);
                    } catch (err) {}
                    try {
                        tempVar143 = cyVar18._getAdjacentRange(null, undefined, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar144 = cyVar18.getRowsBelow(8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar71 = tempVar144;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar145 = cyVar18._handleIdResult(true, 'asf', 3.5, 1, 0, undefined, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar146 = cyVar18.getRowsAbove(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar72 = tempVar146;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar147 = cyVar18.getLastColumn(null, false, false, 0, 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar73 = tempVar147;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar148 = cyVar18.unmerge(null, 34343333, false, 3.5, 34343333, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar149 = cyVar18.getUsedRange('asf', null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar74 = tempVar149;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar150 = cyVar18.track(1, false, 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar151 = cyVar18._ValidateArraySize('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar152 = cyVar18.getEntireColumn(-1, true, 3.5, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar75 = tempVar152;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar153 = cyVar18.getColumnsAfter(null, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar76 = tempVar153;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar154 = cyVar76.getColumnsBefore(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar77 = tempVar154;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar155 = cyVar76.getVisibleView(false, 1, false, true, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar78 = tempVar155;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar156 = cyVar76.getColumnsAfter(0, 34343333, null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar79 = tempVar156;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar157 = cyVar76.unmerge(34343333, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar158 = cyVar76.getEntireColumn(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar80 = tempVar158;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar159 = cyVar76.calculate(34343333, false, 34343333, false, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar160 = cyVar76.getEntireRow(null, null, null, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar81 = tempVar160;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar161 = cyVar76.getRow(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar82 = tempVar161;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar162 = cyVar76.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar83 = tempVar162;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar163 = cyVar76.set(34343333, 1, 1, undefined, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar164 = cyVar76.load(undefined, 0, false, null, undefined, undefined, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar165 = cyVar76.getUsedRangeOrNullObject(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar84 = tempVar165;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar166 = cyVar76.getResizedRange(7, 9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar85 = tempVar166;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar167 = cyVar76.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar86 = tempVar167;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar168 = cyVar76.track(0, 34343333, -1, 1, 'asf', 34343333, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar169 = cyVar76.getUsedRange(3.5, 'asf', 34343333, 1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar87 = tempVar169;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar170 = cyVar76._getAdjacentRange(null, 3.5, 'asf', 1, true, null, -1);
                    } catch (err) {}
                    try {
                        tempVar171 = cyVar76._getAdjacentRange(false, 34343333, 3.5, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar172 = cyVar76.merge('asf', 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar173 = cyVar76.getLastCell(-1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar88 = tempVar173;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar174 = cyVar76.getColumn(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar89 = tempVar174;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar175 = cyVar76.select(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar176 = cyVar76.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar90 = tempVar176;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar177 = cyVar76.getBoundingRect(0, false, 'asf', 34343333, 'asf', 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar91 = tempVar177;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar178 = cyVar76.getRowsBelow(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar92 = tempVar178;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar179 = cyVar76._recursivelySet(-1, 0, false, 0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar180 = cyVar76.getLastColumn(undefined, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar93 = tempVar180;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar181 = cyVar76.getOffsetRange(2, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar94 = tempVar181;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar182 = cyVar76._handleResult(34343333, true, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar183 = cyVar76._handleIdResult(false, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar184 = cyVar76._KeepReference(1, 3.5, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar185 = cyVar76.toJSON('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar95 = tempVar185;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar186 = cyVar76._ensureInteger(null, undefined, -1, false, 3.5);
                    } catch (err) {}
                    try {
                        tempVar187 = cyVar76._ensureInteger(true, null, null, undefined, -1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar188 = cyVar76.getLastRow(true, 'asf', null, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar96 = tempVar188;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar189 = cyVar76.getRowsAbove(-1, false, 34343333, null, false, -1, 34343333, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar97 = tempVar189;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar190 = cyVar76.getCell(1, 11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar98 = tempVar190;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar191 = cyVar75.getLastRow(-1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar99 = tempVar191;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar192 = cyVar75._KeepReference(null, -1, -1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar193 = cyVar75.getIntersectionOrNullObject(34343333, true, true, null, 'asf', 'asf', undefined, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar100 = tempVar193;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar194 = cyVar75._ensureInteger(undefined, 34343333, false, true);
                    } catch (err) {}
                    try {
                        tempVar195 = cyVar75._ensureInteger('asf', 'asf', 1, 'asf', true, false, 34343333, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar196 = cyVar75.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar101 = tempVar196;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar197 = cyVar75.getCell(9, 8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar102 = tempVar197;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar198 = cyVar75._handleIdResult(true, 3.5, 'asf', -1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar199 = cyVar75.getVisibleView(3.5, undefined, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar103 = tempVar199;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar200 = cyVar75.select(1, false, 3.5, 'asf', true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar201 = cyVar75.merge(null, -1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar202 = cyVar75._getAdjacentRange(undefined, undefined, 3.5, false, false, 3.5);
                    } catch (err) {}
                    try {
                        tempVar203 = cyVar75._getAdjacentRange(undefined, 0, 'asf', false, -1, true, 1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar204 = cyVar75.getBoundingRect(3.5, 3.5, 1, false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar104 = tempVar204;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar205 = cyVar75.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar105 = tempVar205;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar206 = cyVar75.calculate('asf', 1, undefined, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar207 = cyVar75.getLastColumn(3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar106 = tempVar207;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar208 = cyVar75._handleResult(1, true, -1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar209 = cyVar75._recursivelySet(3.5, false);
                    } catch (err) {}
                    try {
                        tempVar210 = cyVar75._recursivelySet(3.5, 3.5, 0, false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar211 = cyVar75._ValidateArraySize('asf', false, 3.5, -1, undefined, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar212 = cyVar75.getRowsBelow(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar107 = tempVar212;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar213 = cyVar75.untrack(null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar214 = cyVar75.getEntireRow(undefined, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar108 = tempVar214;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar215 = cyVar75.getRowsAbove('asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar109 = tempVar215;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar216 = cyVar75.getEntireColumn(false, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar110 = tempVar216;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar217 = cyVar75.getResizedRange(false, 3.5, false, 3.5, true, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar111 = tempVar217;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar218 = cyVar75.insert(undefined, true, null, 3.5, 1, 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar112 = tempVar218;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar219 = cyVar75.getUsedRange(3.5, 1, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar113 = tempVar219;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar220 = cyVar75.getColumn(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar114 = tempVar220;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar221 = cyVar75.getLastCell(true, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar115 = tempVar221;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar222 = cyVar75.load(1, 3.5, true, null, 0);
                    } catch (err) {}
                    try {
                        tempVar223 = cyVar75.load(0, false, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar224 = cyVar75.track(1, 34343333, 0, undefined, 0, 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar225 = cyVar75.getOffsetRange(9, 12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar116 = tempVar225;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar226 = cyVar75.unmerge(false, -1, -1, 34343333, null, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar227 = cyVar75.getColumnsAfter(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar117 = tempVar227;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar228 = cyVar75.getColumnsBefore(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar118 = tempVar228;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar229 = cyVar75.getRow(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar119 = tempVar229;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar230 = cyVar75.toJSON(true, 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar120 = tempVar230;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar231 = cyVar74.getColumn(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar121 = tempVar231;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar232 = cyVar74.getCell(12, 5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar122 = tempVar232;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar233 = cyVar74.calculate(-1, false, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar234 = cyVar74.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar123 = tempVar234;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar235 = cyVar74.getResizedRange(7, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar124 = tempVar235;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar236 = cyVar74.getLastCell(1, false, 1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar125 = tempVar236;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar237 = cyVar73._getAdjacentRange(1, true, 'asf', undefined, -1, 'asf', 1, true);
                    } catch (err) {}
                    try {
                        tempVar238 = cyVar73._getAdjacentRange(0, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar239 = cyVar73._ensureInteger(true, 3.5, null, 3.5, -1, 3.5);
                    } catch (err) {}
                    try {
                        tempVar240 = cyVar73._ensureInteger(-1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar241 = cyVar73.getEntireColumn(1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar126 = tempVar241;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar242 = cyVar73.unmerge(null, 0, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar243 = cyVar73.getRowsAbove(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar127 = tempVar243;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar244 = cyVar73.getIntersectionOrNullObject(3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar128 = tempVar244;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar245 = cyVar72._handleIdResult(34343333, undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar246 = cyVar72.getVisibleView(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar129 = tempVar246;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar247 = cyVar72.getIntersection(null, undefined, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar130 = tempVar247;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar248 = cyVar72.getRow(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar131 = tempVar248;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar249 = cyVar72.load(-1);
                    } catch (err) {}
                    try {
                        tempVar250 = cyVar72.load(true, null, 'asf', null, 0, 34343333, 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar251 = cyVar72.set(null, undefined, 'asf', false, true, true);
                    } catch (err) {}
                    try {
                        tempVar252 = cyVar72.set(0, true, undefined, null, true, 'asf', 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar253 = cyVar72._ensureInteger(undefined, null, 0, undefined);
                    } catch (err) {}
                    try {
                        tempVar254 = cyVar72._ensureInteger('asf', 1, true, true, 0, false, 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar255 = cyVar72.toJSON(true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar132 = tempVar255;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar256 = cyVar72.getUsedRange(34343333, null, false, 0, null, 0, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar133 = tempVar256;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar257 = cyVar72.select(undefined, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar258 = cyVar72.track(false, 34343333, undefined, 3.5, 0, 3.5, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar259 = cyVar72.getResizedRange(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar134 = tempVar259;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar260 = cyVar72.getLastRow(null, false, null, false, 1, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar135 = tempVar260;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar261 = cyVar72.getEntireRow(null, false, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar136 = tempVar261;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar262 = cyVar72.getRowsAbove(34343333, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar137 = tempVar262;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar263 = cyVar72.getColumnsAfter(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar138 = tempVar263;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar264 = cyVar72._ValidateArraySize('asf', -1, -1, 0, undefined, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar265 = cyVar72.untrack(-1, 3.5, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar266 = cyVar72._recursivelySet(false, 34343333, true, 1, undefined, true);
                    } catch (err) {}
                    try {
                        tempVar267 = cyVar72._recursivelySet(1, 34343333, 0, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar268 = cyVar72.getOffsetRange(false, true, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar139 = tempVar268;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar269 = cyVar72._KeepReference(1, 1, true, 0, null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar270 = cyVar72.getRowsBelow(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar140 = tempVar270;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar271 = cyVar72.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar141 = tempVar271;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar272 = cyVar72.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar142 = tempVar272;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar273 = cyVar72._getAdjacentRange(3.5, 0);
                    } catch (err) {}
                    try {
                        tempVar274 = cyVar72._getAdjacentRange(1, 34343333, 0, -1, false, null, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar275 = cyVar72.getColumnsBefore(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar143 = tempVar275;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar276 = cyVar72._handleResult(undefined, 1, undefined, 34343333, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar277 = cyVar72.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar144 = tempVar277;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar278 = cyVar72.getColumn(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar145 = tempVar278;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar279 = cyVar72.merge(undefined, true, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar280 = cyVar72.getLastCell(34343333, -1, true, -1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar146 = tempVar280;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar281 = cyVar72.getCell(0, 0, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar147 = tempVar281;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar282 = cyVar72.insert('asf', 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar148 = tempVar282;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar283 = cyVar72.calculate(true, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar284 = cyVar72.getEntireColumn(true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar149 = tempVar284;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar285 = cyVar72.getLastColumn('asf', 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar150 = tempVar285;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar286 = cyVar71.getRow(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar151 = tempVar286;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar287 = cyVar71.getOffsetRange(5, 7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar152 = tempVar287;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar288 = cyVar71._handleResult(34343333, null, undefined, 34343333, null, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar289 = cyVar71.toJSON(-1, 3.5, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar153 = tempVar289;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar290 = cyVar71._handleIdResult(-1, 1, undefined, undefined, -1, 34343333, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar291 = cyVar71.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar154 = tempVar291;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar292 = cyVar69.getOffsetRange(5, 2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar155 = tempVar292;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar293 = cyVar69.merge(34343333, -1, 0, 3.5, 'asf', true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar294 = cyVar69.getEntireRow(null, 3.5, false, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar156 = tempVar294;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar295 = cyVar69.getVisibleView(34343333, 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar157 = tempVar295;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar296 = cyVar69.getCell(undefined, false, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar158 = tempVar296;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar297 = cyVar69.untrack(true, true, null, 3.5, undefined, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar298 = cyVar69.getLastColumn(3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar159 = tempVar298;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar299 = cyVar69.getIntersection(null, 3.5, 0, 'asf', undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar160 = tempVar299;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar300 = cyVar69.calculate(3.5, 34343333, 1, undefined, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar301 = cyVar69.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar161 = tempVar301;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar302 = cyVar69.select('asf', 'asf', 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar303 = cyVar69.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar162 = tempVar303;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar304 = cyVar69.getLastCell(false, null, 'asf', true, 1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar163 = tempVar304;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar305 = cyVar69.getUsedRange(0, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar164 = tempVar305;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar306 = cyVar69.getLastRow(34343333, 1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar165 = tempVar306;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar307 = cyVar69.set(true, 34343333, false, undefined, 3.5, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar308 = cyVar69.getRow(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar166 = tempVar308;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar309 = cyVar69.getRowsBelow(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar167 = tempVar309;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar310 = cyVar69.getEntireColumn(null, true, true, 34343333, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar168 = tempVar310;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar311 = cyVar69._handleIdResult(3.5, null, 34343333, -1, 0, undefined, 1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar312 = cyVar69.getIntersectionOrNullObject(3.5, undefined, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar169 = tempVar312;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar313 = cyVar69._ensureInteger(undefined, 34343333, 0, 1, 0, false);
                    } catch (err) {}
                    try {
                        tempVar314 = cyVar69._ensureInteger(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar315 = cyVar69.unmerge(false, true, true, -1, false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar316 = cyVar69._getAdjacentRange(false, false, 3.5);
                    } catch (err) {}
                    try {
                        tempVar317 = cyVar69._getAdjacentRange(null, null, null, 3.5, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar318 = cyVar69.track(false, 'asf', true, 0, -1, null, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar319 = cyVar69.getRowsAbove(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar170 = tempVar319;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar320 = cyVar69._KeepReference(-1, 34343333, 'asf', undefined, 1, 34343333, 0, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar321 = cyVar69.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar171 = tempVar321;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar322 = cyVar69._recursivelySet(0, 34343333, 'asf', 3.5, 'asf', 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar323 = cyVar69._ValidateArraySize(null, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar324 = cyVar69.load(null, 3.5, undefined, 34343333, undefined);
                    } catch (err) {}
                    try {
                        tempVar325 = cyVar69.load(true, null, null, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar326 = cyVar69._handleResult(true, 0, 0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar327 = cyVar69.getResizedRange(null, 3.5, 34343333, 3.5, null, 0, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar172 = tempVar327;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar328 = cyVar69.getColumnsBefore(false, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar173 = tempVar328;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar329 = cyVar68._handleIdResult(3.5, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar330 = cyVar68.getEntireRow(undefined, false, 1, 1, 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar174 = tempVar330;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar331 = cyVar68.merge(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar332 = cyVar68.getBoundingRect('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar175 = tempVar332;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar333 = cyVar68.getColumn(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar176 = tempVar333;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar334 = cyVar68.getVisibleView(undefined, undefined, 34343333, 3.5, 0, -1, undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar177 = tempVar334;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar335 = cyVar67.getOffsetRange(6, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar178 = tempVar335;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar336 = cyVar67.merge(null, 3.5, false, null, 3.5, -1, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar337 = cyVar67.track(undefined, false, true, false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar338 = cyVar67._ensureInteger(-1, 3.5, 34343333, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar339 = cyVar67.unmerge(1, true, 0, -1, 1, 3.5, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar340 = cyVar67.getRowsBelow(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar179 = tempVar340;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar341 = cyVar67.untrack(-1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar342 = cyVar67.getLastColumn('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar180 = tempVar342;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar343 = cyVar67.getEntireColumn(1, false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar181 = tempVar343;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar344 = cyVar67._recursivelySet(true, null, 'asf');
                    } catch (err) {}
                    try {
                        tempVar345 = cyVar67._recursivelySet(undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar346 = cyVar67.getLastCell('asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar182 = tempVar346;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar347 = cyVar67.calculate(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar348 = cyVar67._ValidateArraySize(false, 0, -1, 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar349 = cyVar67.select(null, 1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar350 = cyVar67.getLastRow(0, true, null, null, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar183 = tempVar350;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar351 = cyVar67.getRowsAbove(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar184 = tempVar351;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar352 = cyVar67.getResizedRange(-1, 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar185 = tempVar352;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar353 = cyVar67._handleIdResult(0, 'asf', 3.5, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar354 = cyVar67.getVisibleView(1, false, 34343333, 'asf', false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar186 = tempVar354;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar355 = cyVar67.set(0, 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar356 = cyVar67.getColumn(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar187 = tempVar356;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar357 = cyVar67.getUsedRange(true, true, 1, 1, 3.5, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar188 = tempVar357;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar358 = cyVar67.getRow(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar189 = tempVar358;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar359 = cyVar67.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar190 = tempVar359;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar360 = cyVar67.load(-1, 0, null, 3.5, null, 1, 34343333, 34343333);
                    } catch (err) {}
                    try {
                        tempVar361 = cyVar67.load(0, 0, false, -1, false, true, null, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar362 = cyVar67._handleResult(undefined, 3.5, undefined, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar363 = cyVar67.getEntireRow(3.5, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar191 = tempVar363;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar364 = cyVar67.getColumnsAfter(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar192 = tempVar364;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar365 = cyVar67._getAdjacentRange(false);
                    } catch (err) {}
                    try {
                        tempVar366 = cyVar67._getAdjacentRange('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar367 = cyVar66.toJSON(true, null, false, undefined, 0, true, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar193 = tempVar367;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar368 = cyVar66.load(1, 1);
                    } catch (err) {}
                    try {
                        tempVar369 = cyVar66.load(null, undefined, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar370 = cyVar66.set('asf', false, 'asf');
                    } catch (err) {}
                    try {
                        tempVar371 = cyVar66.set(3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar372 = cyVar66._handleResult(1, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar373 = cyVar66._handleIdResult(34343333, null, false, null, 1, 0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar374 = cyVar66._recursivelySet(true, true, 1, null, false, null, 34343333, false);
                    } catch (err) {}
                    try {
                        tempVar375 = cyVar66._recursivelySet(-1, undefined, 0, 0, 1, null, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar376 = cyVar65._handleResult(false, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar377 = cyVar65.set(-1, 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar378 = cyVar65.select(null, 'asf', null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar379 = cyVar65._getAdjacentRange(false);
                    } catch (err) {}
                    try {
                        tempVar380 = cyVar65._getAdjacentRange(3.5, -1, null, false, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar381 = cyVar65.getUsedRange(false, 0, 34343333, -1, 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar194 = tempVar381;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar382 = cyVar65.getRowsBelow(undefined, 3.5, undefined, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar195 = tempVar382;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar383 = cyVar65.getLastCell('asf', 34343333, true, null, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar196 = tempVar383;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar384 = cyVar65.track(undefined, 1, undefined, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar385 = cyVar65.getResizedRange(3, 6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar197 = tempVar385;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar386 = cyVar65.getColumnsBefore(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar198 = tempVar386;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar387 = cyVar65._ensureInteger('asf', 'asf', 34343333);
                    } catch (err) {}
                    try {
                        tempVar388 = cyVar65._ensureInteger('asf', 3.5, 1, false, -1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar389 = cyVar65.getEntireColumn(0, -1, 34343333, 1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar199 = tempVar389;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar390 = cyVar65.getRow(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar200 = tempVar390;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar391 = cyVar65.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar201 = tempVar391;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar392 = cyVar65.getBoundingRect(3.5, 34343333, false, 34343333, 1, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar202 = tempVar392;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar393 = cyVar65.getColumnsAfter(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar203 = tempVar393;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar394 = cyVar65.getLastRow(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar204 = tempVar394;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar395 = cyVar65.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar205 = tempVar395;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar396 = cyVar65.getColumn(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar206 = tempVar396;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar397 = cyVar65.toJSON(-1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar207 = tempVar397;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar398 = cyVar65.getRowsAbove(-1, true, 0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar208 = tempVar398;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar399 = cyVar65.getLastColumn(undefined, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar209 = tempVar399;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar400 = cyVar65.unmerge(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar401 = cyVar65.untrack(1, undefined, null, null, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar402 = cyVar65.merge(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar403 = cyVar65.getCell(12, 3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar210 = tempVar403;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar404 = cyVar65.calculate(1, 34343333, true, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar405 = cyVar65._ValidateArraySize(false, undefined, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar406 = cyVar65.getEntireRow(null, 1, 'asf', true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar211 = tempVar406;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar407 = cyVar65.getVisibleView(0, 'asf', -1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar212 = tempVar407;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar408 = cyVar65._KeepReference(false, false, -1, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar409 = cyVar65.load(false, 'asf', false, 3.5, null);
                    } catch (err) {}
                    try {
                        tempVar410 = cyVar65.load(null, undefined, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar411 = cyVar65._recursivelySet(-1, undefined, 0, undefined);
                    } catch (err) {}
                    try {
                        tempVar412 = cyVar65._recursivelySet(null, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar413 = cyVar65.getOffsetRange(12, 3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar213 = tempVar413;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar414 = cyVar65._handleIdResult(null, 0, true, 'asf', 'asf', -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar415 = cyVar64._handleResult(0, null, 1, 'asf', 1, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar416 = cyVar64._recursivelySet(true, 34343333);
                    } catch (err) {}
                    try {
                        tempVar417 = cyVar64._recursivelySet(true, 34343333, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar418 = cyVar64.getLastCell(1, false, 'asf', 0, null, 3.5, 'asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar214 = tempVar418;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar419 = cyVar64.getRow(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar215 = tempVar419;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar420 = cyVar64.track(34343333, true, 'asf', false, true, -1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar421 = cyVar64.getOffsetRange(4, 10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar216 = tempVar421;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar422 = cyVar64.getRowsAbove(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar217 = tempVar422;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar423 = cyVar64.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar218 = tempVar423;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar424 = cyVar64.getEntireRow(-1, null, false, undefined, 34343333, null, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar219 = tempVar424;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar425 = cyVar64.untrack(undefined, 'asf', false, 1, true, undefined, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar426 = cyVar64.insert(-1, 1, -1, null, 3.5, null, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar220 = tempVar426;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar427 = cyVar64.getColumn(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar221 = tempVar427;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar428 = cyVar64.getLastColumn(-1, 0, 'asf', false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar222 = tempVar428;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar429 = cyVar64.unmerge(null, 0, null, 'asf', 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar430 = cyVar64.getCell(6, 9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar223 = tempVar430;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar431 = cyVar64._ValidateArraySize(true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar432 = cyVar64.getColumnsBefore(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar224 = tempVar432;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar433 = cyVar64.set(null, -1, 3.5, 34343333, -1, false);
                    } catch (err) {}
                    try {
                        tempVar434 = cyVar64.set(false, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar435 = cyVar64.calculate(null, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar436 = cyVar64.merge(false, 1, 'asf', true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar437 = cyVar63._handleIdResult('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar438 = cyVar63.load(1, -1, 'asf', -1, true, true, -1, 1);
                    } catch (err) {}
                    try {
                        tempVar439 = cyVar63.load(3.5, 34343333, 34343333, undefined, undefined, undefined, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar440 = cyVar63._recursivelySet(false, 0, undefined, true);
                    } catch (err) {}
                    try {
                        tempVar441 = cyVar63._recursivelySet(undefined, 3.5, true, null, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar442 = cyVar63.getLastCell(false, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar225 = tempVar442;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar443 = cyVar63.toJSON(true, null, 1, 34343333, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar226 = tempVar443;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar444 = cyVar63.getColumnsBefore(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar227 = tempVar444;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar445 = cyVar63.getUsedRange(undefined, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar228 = tempVar445;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar446 = cyVar63.getEntireRow(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar229 = tempVar446;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar447 = cyVar63.getLastRow(34343333, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar230 = tempVar447;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar448 = cyVar63.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar231 = tempVar448;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar449 = cyVar63._ensureInteger(-1, 34343333, 3.5, 34343333, 34343333, 3.5, 34343333, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar450 = cyVar63.getVisibleView(1, 34343333, 1, 'asf', 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar232 = tempVar450;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar451 = cyVar63.unmerge(undefined, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar452 = cyVar63.getUsedRangeOrNullObject('asf', false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar233 = tempVar452;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar453 = cyVar63.getLastColumn(34343333, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar234 = tempVar453;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar454 = cyVar63.getRow(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar235 = tempVar454;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar455 = cyVar63.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar236 = tempVar455;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar456 = cyVar18.getLastColumn(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar237 = tempVar456;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar457 = cyVar18._ensureInteger(false, true, 34343333, 'asf', 1, false, 3.5);
                    } catch (err) {}
                    try {
                        tempVar458 = cyVar18._ensureInteger(-1, 0, 0, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar459 = cyVar18.calculate(undefined, null, 0, true, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar460 = cyVar18.set(1, 'asf', null, 1, 0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar461 = cyVar18.getRowsBelow(8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar238 = tempVar461;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar462 = cyVar18._handleResult(false, true, 3.5, 34343333, 3.5, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar463 = cyVar62.getVisibleView(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar239 = tempVar463;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar464 = cyVar62.toJSON(0, 1, 3.5, 'asf', true, 1, false, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar240 = tempVar464;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar465 = cyVar62.unmerge(34343333, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar466 = cyVar62.getOffsetRange(6, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar241 = tempVar466;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar467 = cyVar62.getEntireRow(34343333, 3.5, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar242 = tempVar467;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar468 = cyVar62._ensureInteger(1, 3.5, 34343333, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar469 = cyVar62.getLastCell(0, -1, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar243 = tempVar469;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar470 = cyVar62.load(3.5, -1, null);
                    } catch (err) {}
                    try {
                        tempVar471 = cyVar62.load(-1, 1, 1, 1, 3.5, 3.5, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar472 = cyVar62.getColumnsAfter(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar244 = tempVar472;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar473 = cyVar62.untrack(0, true, 34343333, 'asf', 34343333, -1, 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar474 = cyVar62.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar245 = tempVar474;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar475 = cyVar61.getColumnsBefore(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar246 = tempVar475;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar476 = cyVar61.getLastColumn(-1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar247 = tempVar476;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar477 = cyVar61.getUsedRange(true, 1, 'asf', true, 1, -1, undefined, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar248 = tempVar477;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar478 = cyVar61.getLastCell(3.5, undefined, 34343333, undefined, 3.5, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar249 = tempVar478;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar479 = cyVar61.getIntersection(-1, 3.5, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar250 = tempVar479;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar480 = cyVar61.calculate(-1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar481 = cyVar60.getUsedRange(false, true, 3.5, 1, 1, 'asf', 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar251 = tempVar481;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar482 = cyVar60.getColumnsAfter(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar252 = tempVar482;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar483 = cyVar60.getBoundingRect(3.5, undefined, 0, undefined, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar253 = tempVar483;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar484 = cyVar60.getLastCell('asf', 1, true, 0, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar254 = tempVar484;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar485 = cyVar60.merge(3.5, undefined, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar486 = cyVar60.unmerge(undefined, 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar487 = cyVar60.select(0, 0, 34343333, null, false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar488 = cyVar60.track(1, 'asf', true, true, -1, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar489 = cyVar60.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar255 = tempVar489;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar490 = cyVar59.calculate(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar491 = cyVar59.getVisibleView('asf', 1, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar256 = tempVar491;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar492 = cyVar59.getEntireRow(3.5, 3.5, 1, false, 'asf', 1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar257 = tempVar492;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar493 = cyVar59._ValidateArraySize(-1, null, false, true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar494 = cyVar59.load(undefined, 3.5, true, 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar495 = cyVar59.track(-1, 'asf', 0, null, 34343333, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar496 = cyVar59.getBoundingRect(null, 1, 1, 'asf', true, 1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar258 = tempVar496;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar497 = cyVar59.getCell(3, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar259 = tempVar497;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar498 = cyVar59.getRow(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar260 = tempVar498;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar499 = cyVar59._getAdjacentRange(true, undefined, false, 3.5, null, 34343333, 'asf');
                    } catch (err) {}
                    try {
                        tempVar500 = cyVar59._getAdjacentRange(false, -1, 34343333, null, 34343333, true, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar501 = cyVar59.select(3.5, 'asf', 'asf', undefined, 34343333, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar502 = cyVar59._KeepReference(undefined, 34343333, true, -1, undefined, false, 3.5, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar503 = cyVar59._handleIdResult('asf', 0, 34343333, null, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar504 = cyVar59.getColumn(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar261 = tempVar504;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar505 = cyVar59._ensureInteger('asf', 0, undefined, -1, undefined, true);
                    } catch (err) {}
                    try {
                        tempVar506 = cyVar59._ensureInteger(3.5, -1, true, -1, 'asf', -1, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar507 = cyVar59.getColumnsBefore(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar262 = tempVar507;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar508 = cyVar59.getRowsBelow(3.5, 3.5, null, undefined, -1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar263 = tempVar508;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar509 = cyVar59.set(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar510 = cyVar59.getRowsAbove(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar264 = tempVar510;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar511 = cyVar59.getLastCell(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar265 = tempVar511;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar512 = cyVar59.unmerge(false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar513 = cyVar59.merge(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar514 = cyVar59.getOffsetRange(1, -1, 34343333, true, 0, 1, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar266 = tempVar514;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar515 = cyVar59.toJSON(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar267 = tempVar515;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar516 = cyVar59.getColumnsAfter(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar268 = tempVar516;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar517 = cyVar59.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar269 = tempVar517;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar518 = cyVar59._handleResult(-1, true, 'asf', 'asf', 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar519 = cyVar59._recursivelySet(34343333, true, false, null, 0);
                    } catch (err) {}
                    try {
                        tempVar520 = cyVar59._recursivelySet(undefined, 3.5, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar521 = cyVar59.getLastColumn(1, null, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar270 = tempVar521;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar522 = cyVar59.getEntireColumn(0, 0, 0, 3.5, false, undefined, 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar271 = tempVar522;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar523 = cyVar58._handleResult(0, 1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar524 = cyVar58.getRow(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar272 = tempVar524;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar525 = cyVar58.calculate(false, -1, 'asf', true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar526 = cyVar58.unmerge(34343333, 'asf', false, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar527 = cyVar58._getAdjacentRange(true);
                    } catch (err) {}
                    try {
                        tempVar528 = cyVar58._getAdjacentRange(0, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar529 = cyVar58._KeepReference(null, true, false, undefined, 'asf', null, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar530 = cyVar58.untrack(null, 3.5, 34343333, 34343333, null, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar531 = cyVar58.track(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar532 = cyVar58.toJSON(undefined, 34343333, 34343333, 1, 'asf', null, -1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar273 = tempVar532;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar533 = cyVar58.getColumn(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar274 = tempVar533;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar534 = cyVar58.getResizedRange(34343333, true, 'asf', true, false, 1, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar275 = tempVar534;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar535 = cyVar58._ensureInteger(true, null, false, undefined, false, 1);
                    } catch (err) {}
                    try {
                        tempVar536 = cyVar58._ensureInteger(0, false, null, false, 34343333, true, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar537 = cyVar58.getOffsetRange(11, 12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar276 = tempVar537;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar538 = cyVar58._recursivelySet(-1, undefined);
                    } catch (err) {}
                    try {
                        tempVar539 = cyVar58._recursivelySet(1, -1, null, 3.5, 3.5, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar540 = cyVar58.getColumnsBefore(1, 1, 0, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar277 = tempVar540;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar541 = cyVar58.getEntireRow(3.5, false, 3.5, 'asf', 'asf', 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar278 = tempVar541;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar542 = cyVar58.getBoundingRect(1, 34343333, 34343333, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar279 = tempVar542;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar543 = cyVar58.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar280 = tempVar543;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar544 = cyVar58.getColumnsAfter(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar281 = tempVar544;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar545 = cyVar58._handleIdResult(-1, false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar546 = cyVar58.getVisibleView(null, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar282 = tempVar546;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar547 = cyVar58.merge(3.5, 'asf', true, 3.5, 34343333, 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar548 = cyVar58.getLastColumn(3.5, 1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar283 = tempVar548;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar549 = cyVar58.getLastCell(null, false, 1, 3.5, true, null, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar284 = tempVar549;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar550 = cyVar57._ValidateArraySize(true, null, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar551 = cyVar57.getResizedRange(13, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar285 = tempVar551;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar552 = cyVar57.load('asf', 34343333, false, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar553 = cyVar57.getRowsBelow(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar286 = tempVar553;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar554 = cyVar57.getLastCell(null, 'asf', 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar287 = tempVar554;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar555 = cyVar57._getAdjacentRange(1, false, 3.5, 34343333, undefined, -1, true, 34343333);
                    } catch (err) {}
                    try {
                        tempVar556 = cyVar57._getAdjacentRange(3.5, true, 'asf', 3.5, 1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar557 = cyVar57.getEntireRow(true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar288 = tempVar557;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar558 = cyVar57.getBoundingRect(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar289 = tempVar558;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar559 = cyVar57.calculate(false, 'asf', false, undefined, 'asf', null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar560 = cyVar57.track(null, 1, true, -1, false, 34343333, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar561 = cyVar57.getRow(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar290 = tempVar561;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar562 = cyVar57.getOffsetRange(8, 11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar291 = tempVar562;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar563 = cyVar57.set(null, false, -1, 0, 'asf', -1);
                    } catch (err) {}
                    try {
                        tempVar564 = cyVar57.set(34343333, -1, 0, 'asf', false, null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar565 = cyVar17._recursivelySet(0, null, 0, false, undefined, null, null, 3.5);
                    } catch (err) {}
                    try {
                        tempVar566 = cyVar17._recursivelySet(34343333, undefined, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar567 = cyVar17._ValidateArraySize(34343333, false, 0, 'asf', 3.5, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar568 = cyVar17.getColumnsAfter(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar292 = tempVar568;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar569 = cyVar17.unmerge(true, 'asf', -1, 1, false, false, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar570 = cyVar17.load(34343333, 34343333, 'asf', true, false, undefined);
                    } catch (err) {}
                    try {
                        tempVar571 = cyVar17.load(1, undefined, false, 34343333, 3.5, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar572 = cyVar17.getColumn(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar293 = tempVar572;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar573 = cyVar17.set('asf', 3.5, false, true, undefined, 34343333);
                    } catch (err) {}
                    try {
                        tempVar574 = cyVar17.set(34343333, true, 34343333, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar575 = cyVar17.track(null, null, null, true, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar576 = cyVar17.getRow(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar294 = tempVar576;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar577 = cyVar17._ensureInteger(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar578 = cyVar17._handleResult(undefined, null, 1, undefined, 0, 0, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar579 = cyVar17._KeepReference(undefined, 'asf', 3.5, 3.5, null, null, false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar580 = cyVar17.getRowsAbove(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar295 = tempVar580;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar581 = cyVar17.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar296 = tempVar581;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar582 = cyVar17.getOffsetRange(8, 8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar297 = tempVar582;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar583 = cyVar17._getAdjacentRange(3.5, 1, undefined, -1, -1, 3.5);
                    } catch (err) {}
                    try {
                        tempVar584 = cyVar17._getAdjacentRange(null, 'asf', false, 1, true, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar585 = cyVar17.insert(null, 3.5, 3.5, undefined, false, -1, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar298 = tempVar585;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar586 = cyVar17.getResizedRange(false, null, false, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar299 = tempVar586;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar587 = cyVar17._handleIdResult(3.5, true, false, null, -1, null, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar588 = cyVar17.getVisibleView(false, 1, true, 3.5, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar300 = tempVar588;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar589 = cyVar17.getEntireColumn(-1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar301 = tempVar589;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar590 = cyVar17.getColumnsBefore(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar302 = tempVar590;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar591 = cyVar17.getRowsBelow(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar303 = tempVar591;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar592 = cyVar303.getLastCell(undefined, 34343333, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar304 = tempVar592;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar593 = cyVar303._handleResult('asf', -1, null, undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar594 = cyVar303.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar305 = tempVar594;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar595 = cyVar303.getRow(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar306 = tempVar595;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar596 = cyVar303.getOffsetRange(5, 9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar307 = tempVar596;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar597 = cyVar303.getColumn(3.5, 'asf', 0, null, null, 0, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar308 = tempVar597;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar598 = cyVar302.load(1, 1, 1, 3.5, 0, 'asf', 0, 1);
                    } catch (err) {}
                    try {
                        tempVar599 = cyVar302.load('asf', 0, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar600 = cyVar302._handleIdResult(3.5, null, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar601 = cyVar302.getOffsetRange(1, 2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar309 = tempVar601;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar602 = cyVar302.untrack(undefined, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar603 = cyVar302.getResizedRange(10, 3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar310 = tempVar603;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar604 = cyVar302.getLastRow(1, true, -1, 0, false, 1, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar311 = tempVar604;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar605 = cyVar302.getColumn(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar312 = tempVar605;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar606 = cyVar302.getLastCell(undefined, false, 34343333, 34343333, null, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar313 = tempVar606;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar607 = cyVar302.getRow(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar314 = tempVar607;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar608 = cyVar302.getIntersectionOrNullObject(1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar315 = tempVar608;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar609 = cyVar302.getEntireRow(3.5, 'asf', 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar316 = tempVar609;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar610 = cyVar302._KeepReference(-1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar611 = cyVar302.getUsedRange(false, 3.5, false, 'asf', 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar317 = tempVar611;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar612 = cyVar302.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar318 = tempVar612;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar613 = cyVar302.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar319 = tempVar613;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar614 = cyVar302.merge(false, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar615 = cyVar302.getCell(2, 8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar320 = tempVar615;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar616 = cyVar302.getLastColumn(undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar321 = tempVar616;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar617 = cyVar302._recursivelySet(false, false, 3.5);
                    } catch (err) {}
                    try {
                        tempVar618 = cyVar302._recursivelySet(true, 3.5, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar619 = cyVar302.calculate(undefined, 34343333, undefined, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar620 = cyVar302.getColumnsAfter(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar322 = tempVar620;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar621 = cyVar302._getAdjacentRange(true);
                    } catch (err) {}
                    try {
                        tempVar622 = cyVar302._getAdjacentRange(null, 34343333, undefined, 3.5, -1, null, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar623 = cyVar302._ValidateArraySize(true, true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar624 = cyVar302.getRowsBelow(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar323 = tempVar624;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar625 = cyVar302.getVisibleView(34343333, true, 'asf', -1, 'asf', 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar324 = tempVar625;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar626 = cyVar302.set(false, 'asf', 1, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar627 = cyVar302.unmerge('asf', 'asf', -1, 34343333, undefined, true, 0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar628 = cyVar302._ensureInteger(undefined, true, null, -1, 3.5, 3.5);
                    } catch (err) {}
                    try {
                        tempVar629 = cyVar302._ensureInteger(0, false, 34343333, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar630 = cyVar301.getRow(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar325 = tempVar630;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar631 = cyVar301.getOffsetRange(13, 6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar326 = tempVar631;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar632 = cyVar301.load(null, -1, 1, 'asf', 1, 3.5, 1);
                    } catch (err) {}
                    try {
                        tempVar633 = cyVar301.load(-1, -1, true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar634 = cyVar301._recursivelySet('asf', null, null);
                    } catch (err) {}
                    try {
                        tempVar635 = cyVar301._recursivelySet(34343333, 0, 1, -1, -1, null, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar636 = cyVar301.getEntireColumn(true, 0, 'asf', 0, -1, false, true, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar327 = tempVar636;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar637 = cyVar301.getColumn(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar328 = tempVar637;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar638 = cyVar301._handleIdResult(3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar639 = cyVar301.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar329 = tempVar639;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar640 = cyVar301.getRowsAbove(1, undefined, true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar330 = tempVar640;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar641 = cyVar301.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar331 = tempVar641;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar642 = cyVar301._handleResult(null, 34343333, true, undefined, undefined, 34343333, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar643 = cyVar301._KeepReference('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar644 = cyVar301.getLastCell(34343333, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar332 = tempVar644;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar645 = cyVar301.getCell(34343333, 0, 34343333, -1, true, 1, true, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar333 = tempVar645;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar646 = cyVar301.unmerge(null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar647 = cyVar301.getUsedRange(-1, 'asf', 3.5, false, 34343333, 'asf', null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar334 = tempVar647;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar648 = cyVar301.getEntireRow(-1, 3.5, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar335 = tempVar648;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar649 = cyVar301.calculate(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar650 = cyVar300.toJSON(true, -1, 'asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar336 = tempVar650;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar651 = cyVar300._handleIdResult(-1, null, undefined, 3.5, 'asf', 3.5, 'asf', 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar652 = cyVar300._recursivelySet(1, true);
                    } catch (err) {}
                    try {
                        tempVar653 = cyVar300._recursivelySet(3.5, 'asf', false, undefined, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar654 = cyVar300._handleResult(true, 1, 'asf', 0, 0, -1, 34343333, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar655 = cyVar300.load(3.5, true);
                    } catch (err) {}
                    try {
                        tempVar656 = cyVar300.load(-1, false, undefined, 'asf', 'asf', 3.5, 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar657 = cyVar300.set(-1, 0, 34343333, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar658 = cyVar299._handleIdResult(-1, null, -1, -1, null, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar659 = cyVar299.getColumnsAfter(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar337 = tempVar659;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar660 = cyVar299._ensureInteger(undefined, 3.5, 'asf', 0, false, 0, true);
                    } catch (err) {}
                    try {
                        tempVar661 = cyVar299._ensureInteger(-1, -1, true, 34343333, true, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar662 = cyVar299.getLastRow(undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar338 = tempVar662;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar663 = cyVar299.load(0);
                    } catch (err) {}
                    try {
                        tempVar664 = cyVar299.load(1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar665 = cyVar299.getIntersection(-1, 0, 34343333, true, -1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar339 = tempVar665;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar666 = cyVar299.getResizedRange(8, 6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar340 = tempVar666;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar667 = cyVar299._recursivelySet(null);
                    } catch (err) {}
                    try {
                        tempVar668 = cyVar299._recursivelySet('asf', null, 3.5, -1, 0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar669 = cyVar299._KeepReference(1, 0, undefined, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar670 = cyVar299.getEntireColumn(1, 3.5, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar341 = tempVar670;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar671 = cyVar299.getColumnsBefore(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar342 = tempVar671;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar672 = cyVar299.toJSON(-1, 34343333, 34343333, undefined, true, true, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar343 = tempVar672;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar673 = cyVar299.getVisibleView(null, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar344 = tempVar673;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar674 = cyVar299.getRowsBelow(undefined, -1, 3.5, true, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar345 = tempVar674;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar675 = cyVar299.getRow(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar346 = tempVar675;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar676 = cyVar299.track(null, -1, -1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar677 = cyVar299.getLastColumn(true, false, null, 0, false, 1, 'asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar347 = tempVar677;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar678 = cyVar299.getCell(11, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar348 = tempVar678;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar679 = cyVar299.unmerge(undefined, 'asf', undefined, null, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar680 = cyVar299.getColumn(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar349 = tempVar680;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar681 = cyVar299.insert(undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar350 = tempVar681;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar682 = cyVar299.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar351 = tempVar682;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar683 = cyVar299._handleResult(undefined, 3.5, false, false, null, 34343333, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar684 = cyVar299.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar352 = tempVar684;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar685 = cyVar299.merge(0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar686 = cyVar299.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar353 = tempVar686;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar687 = cyVar299.getLastCell(-1, 3.5, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar354 = tempVar687;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar688 = cyVar299.getEntireRow(false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar355 = tempVar688;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar689 = cyVar299.select(0, -1, 3.5, undefined, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar690 = cyVar299.set(0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar691 = cyVar299.getUsedRange(null, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar356 = tempVar691;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar692 = cyVar298.merge(false, 34343333, 0, true, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar693 = cyVar298.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar357 = tempVar693;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar694 = cyVar298.track(false, undefined, -1, undefined, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar695 = cyVar298.getOffsetRange(6, 2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar358 = tempVar695;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar696 = cyVar298.untrack(undefined, false, 0, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar697 = cyVar298.getLastCell('asf', null, undefined, undefined, 0, null, 34343333, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar359 = tempVar697;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar698 = cyVar298._handleIdResult('asf', 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar699 = cyVar298.calculate(3.5, true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar700 = cyVar298.getLastColumn(0, null, 0, 'asf', -1, -1, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar360 = tempVar700;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar701 = cyVar298.getColumn(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar361 = tempVar701;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar702 = cyVar298.getUsedRange('asf', undefined, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar362 = tempVar702;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar703 = cyVar298._handleResult(3.5, 3.5, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar704 = cyVar298.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar363 = tempVar704;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar705 = cyVar298.getResizedRange(6, 8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar364 = tempVar705;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar706 = cyVar297.set(-1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar707 = cyVar297._getAdjacentRange(undefined, -1, true, -1, 'asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar708 = cyVar297.getColumn(true, false, false, false, -1, 0, 'asf', 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar365 = tempVar708;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar709 = cyVar297.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar366 = tempVar709;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar710 = cyVar297._ValidateArraySize(0, undefined, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar711 = cyVar297.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar367 = tempVar711;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar712 = cyVar297.getUsedRange(3.5, 'asf', 3.5, true, 3.5, 3.5, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar368 = tempVar712;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar713 = cyVar297.select('asf', -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar714 = cyVar297.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar369 = tempVar714;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar715 = cyVar297.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar370 = tempVar715;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar716 = cyVar297.getEntireColumn(false, 34343333, 3.5, null, true, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar371 = tempVar716;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar717 = cyVar297.getCell(-1, 3.5, -1, 'asf', -1, null, 3.5, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar372 = tempVar717;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar718 = cyVar297.getColumnsBefore(false, 'asf', 0, 34343333, 3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar373 = tempVar718;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar719 = cyVar297.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar374 = tempVar719;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar720 = cyVar297._handleIdResult(3.5, 34343333, undefined, null, true, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar721 = cyVar297.unmerge(true, true, 'asf', true, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar722 = cyVar296.toJSON('asf', 'asf', -1, 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar375 = tempVar722;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar723 = cyVar296.getLastColumn(null, -1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar376 = tempVar723;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar724 = cyVar296._KeepReference(3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar725 = cyVar296.calculate('asf', 3.5, 0, null, -1, 34343333, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar726 = cyVar296._handleResult(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar727 = cyVar296.getEntireRow(undefined, 1, true, 'asf', undefined, null, 1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar377 = tempVar727;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar728 = cyVar296._getAdjacentRange(true, undefined, 0, 1, true, 0, 'asf', null);
                    } catch (err) {}
                    try {
                        tempVar729 = cyVar296._getAdjacentRange(-1, -1, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar730 = cyVar296.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar378 = tempVar730;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar731 = cyVar296._ValidateArraySize(false, 1, 1, -1, -1, 34343333, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar732 = cyVar296.getUsedRange(null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar379 = tempVar732;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar733 = cyVar296.set(true, undefined, 34343333, undefined, null, false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar734 = cyVar296.getOffsetRange(6, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar380 = tempVar734;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar735 = cyVar296.getBoundingRect(null, null, 34343333, null, undefined, 34343333, null, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar381 = tempVar735;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar736 = cyVar296.merge(undefined, 1, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar737 = cyVar296.getRow(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar382 = tempVar737;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar738 = cyVar296.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar383 = tempVar738;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar739 = cyVar296.getVisibleView(1, 34343333, null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar384 = tempVar739;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar740 = cyVar296.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar385 = tempVar740;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar741 = cyVar296.getRowsAbove(1, 3.5, -1, 3.5, 3.5, false, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar386 = tempVar741;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar742 = cyVar296.getColumn(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar387 = tempVar742;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar743 = cyVar296.getResizedRange(11, 6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar388 = tempVar743;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar744 = cyVar296.untrack('asf', 3.5, -1, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar745 = cyVar296._recursivelySet(-1, -1, true, false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar746 = cyVar296.track(1, 0, false, true, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar747 = cyVar296.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar389 = tempVar747;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar748 = cyVar296._handleIdResult('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar749 = cyVar296.getLastRow(false, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar390 = tempVar749;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar750 = cyVar296.getLastCell(0, -1, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar391 = tempVar750;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar751 = cyVar296.select('asf', 0, 3.5, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar752 = cyVar296.getCell(1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar392 = tempVar752;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar753 = cyVar296.getRowsBelow(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar393 = tempVar753;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar754 = cyVar295._recursivelySet(34343333);
                    } catch (err) {}
                    try {
                        tempVar755 = cyVar295._recursivelySet(-1, 1, true, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar756 = cyVar295.track(false, null, 34343333, null, 34343333, true, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar757 = cyVar295.getCell(8, 6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar394 = tempVar757;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar758 = cyVar295.getRowsAbove(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar395 = tempVar758;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar759 = cyVar295.getResizedRange(4, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar396 = tempVar759;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar760 = cyVar295.load(null, 34343333, null, 34343333);
                    } catch (err) {}
                    try {
                        tempVar761 = cyVar295.load(3.5, 34343333, 3.5, -1, undefined, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar762 = cyVar295.untrack('asf', 3.5, -1, 1, 3.5, 34343333, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar763 = cyVar295.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar397 = tempVar763;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar764 = cyVar295._handleResult(-1, undefined, 3.5, null, null, -1, 0, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar765 = cyVar295.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar398 = tempVar765;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar766 = cyVar295.getBoundingRect(34343333, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar399 = tempVar766;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar767 = cyVar295.getVisibleView(true, 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar400 = tempVar767;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar768 = cyVar295._handleIdResult(34343333, 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar769 = cyVar295.getEntireColumn(0, -1, 34343333, 'asf', 1, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar401 = tempVar769;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar770 = cyVar295._ensureInteger(34343333, -1, 3.5, -1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar771 = cyVar295._getAdjacentRange(null, 1, 34343333);
                    } catch (err) {}
                    try {
                        tempVar772 = cyVar295._getAdjacentRange(null, 1, 'asf', 3.5, null, 1, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar773 = cyVar295.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar402 = tempVar773;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar774 = cyVar295.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar403 = tempVar774;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar775 = cyVar295.calculate(undefined, 0, false, null, -1, 1, 'asf', null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar776 = cyVar295.getLastCell(3.5, true, null, undefined, 3.5, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar404 = tempVar776;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar777 = cyVar295._ValidateArraySize(3.5, 34343333, false, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar778 = cyVar295.getColumnsBefore(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar405 = tempVar778;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar779 = cyVar295.getColumn(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar406 = tempVar779;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar780 = cyVar295._KeepReference(3.5, 1, 34343333, 34343333, 1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar781 = cyVar294.calculate(34343333, 34343333, false, 1, 'asf', 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar782 = cyVar294.getRowsBelow(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar407 = tempVar782;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar783 = cyVar294.getLastRow(34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar408 = tempVar783;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar784 = cyVar294.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar409 = tempVar784;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar785 = cyVar294.getResizedRange(10, 7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar410 = tempVar785;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar786 = cyVar294.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar411 = tempVar786;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar787 = cyVar17.calculate(undefined, 0, true, undefined, undefined, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar788 = cyVar17.select(null, null, 'asf', 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar789 = cyVar17._handleIdResult(1, 34343333, 'asf', 3.5, 0, 'asf', true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar790 = cyVar17.getColumnsAfter(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar412 = tempVar790;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar791 = cyVar17.getIntersection(3.5, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar413 = tempVar791;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar792 = cyVar17.merge(0, 1, null, true, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar793 = cyVar293.getRowsBelow(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar414 = tempVar793;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar794 = cyVar293.getLastColumn('asf', -1, null, 0, -1, 0, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar415 = tempVar794;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar795 = cyVar293.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar416 = tempVar795;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar796 = cyVar293.untrack(true, 0, null, null, 0, -1, false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar797 = cyVar293.getColumnsBefore(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar417 = tempVar797;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar798 = cyVar293.load(undefined, null, false, 3.5, 34343333, 3.5, true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar799 = cyVar293.getColumnsAfter(null, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar418 = tempVar799;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar800 = cyVar293.calculate(undefined, 3.5, null, 1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar801 = cyVar293._handleIdResult(false, true, 1, null, true, true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar802 = cyVar293.merge(0, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar803 = cyVar293.getCell(7, 3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar419 = tempVar803;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar804 = cyVar293.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar420 = tempVar804;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar805 = cyVar293.getResizedRange(7, 12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar421 = tempVar805;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar806 = cyVar293.select(null, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar807 = cyVar293.getUsedRange(3.5, 1, undefined, 'asf', 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar422 = tempVar807;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar808 = cyVar293.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar423 = tempVar808;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar809 = cyVar293._KeepReference(1, false, 0, 0, 34343333, 0, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar810 = cyVar293.getLastRow('asf', false, 3.5, true, true, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar424 = tempVar810;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar811 = cyVar292.getEntireColumn(true, 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar425 = tempVar811;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar812 = cyVar292.getOffsetRange(11, 7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar426 = tempVar812;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar813 = cyVar292.getColumnsBefore(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar427 = tempVar813;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar814 = cyVar292._handleResult(34343333, false, -1, false, -1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar815 = cyVar292.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar428 = tempVar815;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar816 = cyVar292.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar429 = tempVar816;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar817 = cyVar292.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar430 = tempVar817;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar818 = cyVar292.load(true);
                    } catch (err) {}
                    try {
                        tempVar819 = cyVar292.load(1, 34343333, -1, true, -1, undefined, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar820 = cyVar292._recursivelySet(undefined, undefined, undefined);
                    } catch (err) {}
                    try {
                        tempVar821 = cyVar292._recursivelySet(true, -1, 0, false, 3.5, 34343333, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar822 = cyVar292.select(3.5, 'asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar823 = cyVar292.calculate(0, 1, 0, true, undefined, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar824 = cyVar292.getLastRow(undefined, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar431 = tempVar824;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar825 = cyVar292.track(null, null, -1, -1, true, 3.5, undefined, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar826 = cyVar292.getRow(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar432 = tempVar826;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar827 = cyVar292._handleIdResult(1, null, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar828 = cyVar292.getColumnsAfter(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar433 = tempVar828;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar829 = cyVar292.getLastCell(false, 1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar434 = tempVar829;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar830 = cyVar292.getResizedRange(10, 2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar435 = tempVar830;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar831 = cyVar292.getEntireRow(false, false, false, 1, true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar436 = tempVar831;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar832 = cyVar292.getLastColumn(null, null, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar437 = tempVar832;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar833 = cyVar292.unmerge('asf', 3.5, undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar834 = cyVar16.getOffsetRange(7, 10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar438 = tempVar834;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar835 = cyVar16.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar439 = tempVar835;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar836 = cyVar16.toJSON(-1, true, -1, true, 'asf', 1, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar440 = tempVar836;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar837 = cyVar16.select(false, 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar838 = cyVar16.getLastCell(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar441 = tempVar838;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar839 = cyVar16.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar442 = tempVar839;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar840 = cyVar442.select(3.5, 0, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar841 = cyVar442.unmerge(1, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar842 = cyVar442._KeepReference(0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar843 = cyVar442.getColumnsBefore(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar443 = tempVar843;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar844 = cyVar442.insert(undefined, false, 3.5, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar444 = tempVar844;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar845 = cyVar442._ValidateArraySize('asf', 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar846 = cyVar442.getVisibleView(-1, null, null, -1, -1, 1, 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar445 = tempVar846;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar847 = cyVar442.getRowsAbove(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar446 = tempVar847;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar848 = cyVar442.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar447 = tempVar848;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar849 = cyVar442.load(34343333, true);
                    } catch (err) {}
                    try {
                        tempVar850 = cyVar442.load(false, null, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar851 = cyVar441.getVisibleView(false, 1, null, null, 'asf', 3.5, 34343333, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar448 = tempVar851;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar852 = cyVar441.untrack(null, 34343333, undefined, 0, undefined, null, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar853 = cyVar441.toJSON(0, 1, 3.5, true, false, 3.5, false, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar449 = tempVar853;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar854 = cyVar441.track(true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar855 = cyVar441.getEntireColumn('asf', -1, 1, 1, true, false, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar450 = tempVar855;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar856 = cyVar441.load(1, undefined, null, -1, null, 34343333, 3.5);
                    } catch (err) {}
                    try {
                        tempVar857 = cyVar441.load(0, 1, false, 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar858 = cyVar441.getRowsAbove(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar451 = tempVar858;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar859 = cyVar441._recursivelySet(3.5, 1);
                    } catch (err) {}
                    try {
                        tempVar860 = cyVar441._recursivelySet(-1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar861 = cyVar441.getColumn(true, 1, 1, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar452 = tempVar861;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar862 = cyVar441.getEntireRow(undefined, undefined, null, false, -1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar453 = tempVar862;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar863 = cyVar441.getLastCell(false, 1, 1, null, 1, 'asf', 34343333, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar454 = tempVar863;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar864 = cyVar441.getResizedRange(11, 10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar455 = tempVar864;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar865 = cyVar441._ensureInteger('asf', undefined, false);
                    } catch (err) {}
                    try {
                        tempVar866 = cyVar441._ensureInteger(undefined, 34343333, false, true, -1, 1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar867 = cyVar441.select('asf', 3.5, 0, false, false, true, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar868 = cyVar441.getUsedRangeOrNullObject(34343333, 0, 3.5, 34343333, 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar456 = tempVar868;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar869 = cyVar441.set(true, -1, true, 34343333, undefined, 3.5, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar870 = cyVar441.getOffsetRange(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar457 = tempVar870;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar871 = cyVar441.getUsedRange(undefined, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar458 = tempVar871;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar872 = cyVar441.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar459 = tempVar872;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar873 = cyVar441.getLastRow(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar460 = tempVar873;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar874 = cyVar441.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar461 = tempVar874;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar875 = cyVar441.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar462 = tempVar875;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar876 = cyVar441.getLastColumn(true, false, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar463 = tempVar876;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar877 = cyVar441.unmerge(undefined, 3.5, 3.5, true, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar878 = cyVar439.set(null, 3.5, 3.5, 1, null, 34343333, 3.5, -1);
                    } catch (err) {}
                    try {
                        tempVar879 = cyVar439.set(0, null, undefined, undefined, 34343333, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar880 = cyVar439.getEntireColumn('asf', 3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar464 = tempVar880;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar881 = cyVar439.getVisibleView(true, 'asf', -1, 1, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar465 = tempVar881;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar882 = cyVar439.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar466 = tempVar882;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar883 = cyVar439.getIntersectionOrNullObject(0, 'asf', -1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar467 = tempVar883;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar884 = cyVar439.getColumn(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar468 = tempVar884;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar885 = cyVar439.untrack(1, 34343333, 'asf', undefined, 0, false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar886 = cyVar439.track(0, 34343333, 0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar887 = cyVar439.select(3.5, 'asf', false, 0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar888 = cyVar439._handleResult('asf', 3.5, 'asf', 34343333, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar889 = cyVar438.getColumnsAfter(34343333, true, null, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar469 = tempVar889;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar890 = cyVar438._handleResult(true, 34343333, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar891 = cyVar438.untrack(-1, 3.5, -1, 3.5, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar892 = cyVar438.getRowsAbove(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar470 = tempVar892;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar893 = cyVar438._KeepReference(1, null, 'asf', null, 'asf', -1, -1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar894 = cyVar438._ValidateArraySize(false, 34343333, 3.5, 34343333, undefined, 0, 34343333, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar895 = cyVar438.getUsedRange(-1, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar471 = tempVar895;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar896 = cyVar438.getLastRow(3.5, 'asf', undefined, undefined, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar472 = tempVar896;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar897 = cyVar438.getVisibleView(-1, -1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar473 = tempVar897;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar898 = cyVar438._getAdjacentRange(1, null, undefined, -1, 'asf', 'asf', undefined);
                    } catch (err) {}
                    try {
                        tempVar899 = cyVar438._getAdjacentRange(-1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar900 = cyVar438.set(null, 34343333, undefined, 0, 0);
                    } catch (err) {}
                    try {
                        tempVar901 = cyVar438.set('asf', null, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar902 = cyVar438.getResizedRange(8, 6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar474 = tempVar902;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar903 = cyVar438._handleIdResult(0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar904 = cyVar15.toJSON('asf', 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar475 = tempVar904;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar905 = cyVar15._recursivelySet(false, null, 'asf', undefined, true, 0, true, 1);
                    } catch (err) {}
                    try {
                        tempVar906 = cyVar15._recursivelySet(true, 'asf', -1, 'asf', true, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar907 = cyVar15._handleIdResult(null, -1, 3.5, undefined, 1, 0, undefined, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar908 = cyVar15.set(34343333, false, 3.5, null, false, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar909 = cyVar15.load(3.5, 'asf', 'asf');
                    } catch (err) {}
                    try {
                        tempVar910 = cyVar15.load(1, 0, 'asf', 3.5, 1, false, undefined, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar911 = cyVar15.getRange(1, true, 3.5, 3.5, null, false, false, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar476 = tempVar911;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar912 = cyVar476._handleResult(34343333, 34343333, 'asf', 3.5, 3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar913 = cyVar476._recursivelySet(34343333, -1, 1, 34343333);
                    } catch (err) {}
                    try {
                        tempVar914 = cyVar476._recursivelySet(0, 'asf', 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar915 = cyVar476.getResizedRange(6, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar477 = tempVar915;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar916 = cyVar476._KeepReference(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar917 = cyVar476.calculate(0, false, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar918 = cyVar476._ensureInteger(34343333, 3.5, 3.5, -1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar919 = cyVar14._KeepReference('asf', 3.5, 3.5, 3.5, 0, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar920 = cyVar14.getResizedRange(12, 13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar478 = tempVar920;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar921 = cyVar14._handleResult(-1, null, 'asf', 'asf', true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar922 = cyVar14.getRowsBelow(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar479 = tempVar922;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar923 = cyVar14.getColumnsAfter(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar480 = tempVar923;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar924 = cyVar14.getOffsetRange(8, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar481 = tempVar924;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar925 = cyVar481.merge(false, true, 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar926 = cyVar481._recursivelySet(undefined, -1, 'asf', false);
                    } catch (err) {}
                    try {
                        tempVar927 = cyVar481._recursivelySet(true, true, 3.5, 0, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar928 = cyVar481._getAdjacentRange(0, 34343333);
                    } catch (err) {}
                    try {
                        tempVar929 = cyVar481._getAdjacentRange(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar930 = cyVar481.load(undefined, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar931 = cyVar481.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar482 = tempVar931;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar932 = cyVar481.getLastColumn(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar483 = tempVar932;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar933 = cyVar481._handleIdResult(-1, null, 'asf', 0, undefined, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar934 = cyVar481.getUsedRange(true, 0, -1, null, 1, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar484 = tempVar934;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar935 = cyVar481.unmerge(true, 3.5, -1, -1, undefined, true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar936 = cyVar481._ensureInteger(-1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar937 = cyVar481.getIntersectionOrNullObject('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar485 = tempVar937;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar938 = cyVar481.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar486 = tempVar938;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar939 = cyVar481.getRowsBelow(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar487 = tempVar939;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar940 = cyVar481.getOffsetRange(8, 11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar488 = tempVar940;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar941 = cyVar481._handleResult(undefined, -1, 1, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar942 = cyVar481.select(null, true, 34343333, true, false, undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar943 = cyVar481.set(34343333, -1, null, 34343333, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar944 = cyVar481._KeepReference(3.5, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar945 = cyVar481.getEntireColumn(undefined, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar489 = tempVar945;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar946 = cyVar481.getLastCell(true, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar490 = tempVar946;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar947 = cyVar481.getEntireRow(undefined, true, false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar491 = tempVar947;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar948 = cyVar481.getVisibleView(true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar492 = tempVar948;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar949 = cyVar481.getRowsAbove(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar493 = tempVar949;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar950 = cyVar481.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar494 = tempVar950;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar951 = cyVar481.getColumnsAfter(false, null, false, true, 3.5, undefined, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar495 = tempVar951;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar952 = cyVar480.getLastRow(34343333, 3.5, 3.5, 3.5, 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar496 = tempVar952;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar953 = cyVar480.set(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar954 = cyVar480.getCell(3, 5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar497 = tempVar954;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar955 = cyVar480.getOffsetRange(3, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar498 = tempVar955;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar956 = cyVar480._handleIdResult(false, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar957 = cyVar480._handleResult(true, 34343333, null, 'asf', 1, 0, true, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar958 = cyVar480.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar499 = tempVar958;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar959 = cyVar480.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar500 = tempVar959;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar960 = cyVar480.getRowsBelow(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar501 = tempVar960;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar961 = cyVar480.getColumn(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar502 = tempVar961;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar962 = cyVar480.getRow(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar503 = tempVar962;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar963 = cyVar480.getColumnsBefore(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar504 = tempVar963;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar964 = cyVar480.getVisibleView(false, 34343333, true, 34343333, -1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar505 = tempVar964;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar965 = cyVar480.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar506 = tempVar965;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar966 = cyVar480.getColumnsAfter(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar507 = tempVar966;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar967 = cyVar480.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar508 = tempVar967;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar968 = cyVar480.getLastCell(false, 'asf', 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar509 = tempVar968;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar969 = cyVar480.untrack(undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar970 = cyVar479.unmerge(undefined, 1, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar971 = cyVar479._ensureInteger(3.5, 'asf', false, 0, true, true, -1, 3.5);
                    } catch (err) {}
                    try {
                        tempVar972 = cyVar479._ensureInteger(-1, false, 0, 1, 1, 0, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar973 = cyVar479.getBoundingRect('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar510 = tempVar973;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar974 = cyVar479.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar511 = tempVar974;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar975 = cyVar479.getResizedRange(9, 10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar512 = tempVar975;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar976 = cyVar479.getUsedRange(undefined, 0, undefined, 1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar513 = tempVar976;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar977 = cyVar479._KeepReference(false, -1, false, 1, undefined, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar978 = cyVar479.getColumnsBefore(true, -1, -1, 3.5, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar514 = tempVar978;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar979 = cyVar479.getColumnsAfter(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar515 = tempVar979;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar980 = cyVar479.getRowsAbove(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar516 = tempVar980;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar981 = cyVar479.getLastCell(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar517 = tempVar981;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar982 = cyVar479.untrack(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar983 = cyVar479.getCell(6, 13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar518 = tempVar983;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar984 = cyVar479.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar519 = tempVar984;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar985 = cyVar479.getRowsBelow(3.5, true, 3.5, 3.5, undefined, 0, undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar520 = tempVar985;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar986 = cyVar479.merge(34343333, undefined, 3.5, 34343333, true, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar987 = cyVar479.calculate(-1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar988 = cyVar478._getAdjacentRange(1, 'asf', 'asf', true, undefined, -1, null, false);
                    } catch (err) {}
                    try {
                        tempVar989 = cyVar478._getAdjacentRange('asf', false, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar990 = cyVar478._recursivelySet(null, false, 34343333, undefined, 34343333, 'asf');
                    } catch (err) {}
                    try {
                        tempVar991 = cyVar478._recursivelySet(34343333, 34343333, 0, 3.5, 0, true, -1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar992 = cyVar478._handleIdResult(undefined, false, 3.5, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar993 = cyVar478._ValidateArraySize(3.5, 3.5, 'asf', false, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar994 = cyVar478.untrack(1, 1, null, 3.5, 1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar995 = cyVar478.getLastRow(true, -1, null, -1, 0, null, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar521 = tempVar995;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar996 = cyVar478._KeepReference(null, false, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar997 = cyVar478.getOffsetRange(10, 8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar522 = tempVar997;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar998 = cyVar478.getColumnsBefore(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar523 = tempVar998;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar999 = cyVar478.getResizedRange(13, 3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar524 = tempVar999;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1000 = cyVar478.calculate(0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1001 = cyVar478.getVisibleView('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar525 = tempVar1001;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1002 = cyVar478._ensureInteger(null, 34343333, 0, 0, 34343333, 0);
                    } catch (err) {}
                    try {
                        tempVar1003 = cyVar478._ensureInteger('asf', null, 1, null, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1004 = cyVar478.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar526 = tempVar1004;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1005 = cyVar478.select(1, true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1006 = cyVar478.toJSON('asf', -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar527 = tempVar1006;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1007 = cyVar478.getRowsAbove(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar528 = tempVar1007;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1008 = cyVar478.getColumn(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar529 = tempVar1008;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1009 = cyVar478.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar530 = tempVar1009;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1010 = cyVar478.getLastColumn(-1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar531 = tempVar1010;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1011 = cyVar478.merge(3.5, 0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1012 = cyVar478.unmerge('asf', true, 0, 'asf', 0, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1013 = cyVar478.set(null);
                    } catch (err) {}
                    try {
                        tempVar1014 = cyVar478.set(3.5, 1, undefined, 3.5, true, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1015 = cyVar478.getColumnsAfter(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar532 = tempVar1015;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1016 = cyVar13.toJSON(false, undefined, undefined, undefined, 0, 0, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar533 = tempVar1016;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1017 = cyVar13.getColumn(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar534 = tempVar1017;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1018 = cyVar13.getColumnsAfter(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar535 = tempVar1018;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1019 = cyVar13.getLastCell(0, true, undefined, undefined, true, 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar536 = tempVar1019;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1020 = cyVar13.getLastRow(34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar537 = tempVar1020;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1021 = cyVar13._handleResult(false, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1022 = cyVar13.getEntireColumn('asf', 3.5, null, -1, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar538 = tempVar1022;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1023 = cyVar13.getRowsBelow(1, true, 'asf', 34343333, null, 1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar539 = tempVar1023;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1024 = cyVar13.getRowsAbove(null, undefined, 3.5, 1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar540 = tempVar1024;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1025 = cyVar540.getLastCell(undefined, 34343333, true, 'asf', undefined, 34343333, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar541 = tempVar1025;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1026 = cyVar540.select(0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1027 = cyVar540.load(false, false, 34343333, null, 0, 3.5, 0, null);
                    } catch (err) {}
                    try {
                        tempVar1028 = cyVar540.load(-1, 1, undefined, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1029 = cyVar540._getAdjacentRange(false, null, null, 34343333);
                    } catch (err) {}
                    try {
                        tempVar1030 = cyVar540._getAdjacentRange(3.5, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1031 = cyVar540.getColumn(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar542 = tempVar1031;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1032 = cyVar540.insert(null, true, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar543 = tempVar1032;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1033 = cyVar539.insert(34343333, 1, null, 0, true, null, 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar544 = tempVar1033;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1034 = cyVar539.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar545 = tempVar1034;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1035 = cyVar539._getAdjacentRange(3.5, null, -1, undefined);
                    } catch (err) {}
                    try {
                        tempVar1036 = cyVar539._getAdjacentRange(3.5, null, null, 0, 34343333, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1037 = cyVar539.toJSON(34343333, 3.5, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar546 = tempVar1037;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1038 = cyVar539._KeepReference(0, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1039 = cyVar539._handleIdResult(null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1040 = cyVar538.set(34343333, -1, true, undefined, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1041 = cyVar538._handleIdResult(undefined, 1, 1, 0, 0, true, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1042 = cyVar538._getAdjacentRange(3.5);
                    } catch (err) {}
                    try {
                        tempVar1043 = cyVar538._getAdjacentRange(true, true, 3.5, true, 'asf', 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1044 = cyVar538.getEntireColumn(0, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar547 = tempVar1044;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1045 = cyVar538.getColumn(10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar548 = tempVar1045;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1046 = cyVar538.getEntireRow(false, 1, 1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar549 = tempVar1046;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1047 = cyVar538.getLastRow(3.5, 0, true, 34343333, true, 34343333, false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar550 = tempVar1047;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1048 = cyVar538.getResizedRange(8, 2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar551 = tempVar1048;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1049 = cyVar538.unmerge(34343333, null, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1050 = cyVar538.getRow(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar552 = tempVar1050;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1051 = cyVar538._recursivelySet(true, -1, true, true, false, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1052 = cyVar538.calculate(null, 3.5, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1053 = cyVar538._ensureInteger(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1054 = cyVar538.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar553 = tempVar1054;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1055 = cyVar538.load(34343333, undefined, true, false, false, 0, true, false);
                    } catch (err) {}
                    try {
                        tempVar1056 = cyVar538.load(-1, false, 0, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1057 = cyVar538.getLastColumn(true, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar554 = tempVar1057;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1058 = cyVar538.getColumnsBefore(0, 'asf', 0, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar555 = tempVar1058;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1059 = cyVar538.getBoundingRect(0, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar556 = tempVar1059;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1060 = cyVar538.getCell(9, 11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar557 = tempVar1060;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1061 = cyVar538.getVisibleView(0, false, 34343333, 1, null, 'asf', 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar558 = tempVar1061;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1062 = cyVar538.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar559 = tempVar1062;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1063 = cyVar538.select(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1064 = cyVar538.getRowsBelow(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar560 = tempVar1064;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1065 = cyVar538._ValidateArraySize(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1066 = cyVar538.getColumnsAfter(4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar561 = tempVar1066;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1067 = cyVar538.track('asf', 'asf', 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1068 = cyVar538.getLastCell(false, 'asf', 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar562 = tempVar1068;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1069 = cyVar538.getOffsetRange(8, 8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar563 = tempVar1069;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1070 = cyVar538.untrack(undefined, 3.5, false, 'asf', true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1071 = cyVar538.merge(0, undefined, 34343333, 34343333, 34343333, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1072 = cyVar538.getUsedRange(null, 0, null, null, 34343333, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar564 = tempVar1072;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1073 = cyVar538.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar565 = tempVar1073;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1074 = cyVar538.getRowsAbove(8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar566 = tempVar1074;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1075 = cyVar538._KeepReference(3.5, -1, 3.5, -1, 0, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1076 = cyVar538._handleResult('asf', 3.5, 34343333, 1, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1077 = cyVar538.toJSON(34343333, 0, 'asf', -1, -1, 'asf', false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar567 = tempVar1077;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1078 = cyVar537.insert(3.5, false, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar568 = tempVar1078;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1079 = cyVar537.track(1, 3.5, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1080 = cyVar537._handleIdResult(-1, null, 3.5, 3.5, 0, 3.5, 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1081 = cyVar537.getResizedRange(1, 'asf', true, null, 1, 'asf', -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar569 = tempVar1081;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1082 = cyVar537.merge(true, 'asf', undefined, 3.5, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1083 = cyVar537.untrack(1, 1, -1, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1084 = cyVar537.getRow(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar570 = tempVar1084;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1085 = cyVar537.getUsedRange(3.5, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar571 = tempVar1085;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1086 = cyVar537.getColumnsBefore(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar572 = tempVar1086;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1087 = cyVar537.getRowsAbove(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar573 = tempVar1087;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1088 = cyVar537.getLastCell(true, undefined, 3.5, 1, true, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar574 = tempVar1088;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1089 = cyVar537.getColumnsAfter(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar575 = tempVar1089;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1090 = cyVar537.getIntersection(34343333, true, -1, -1, 34343333, false, 1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar576 = tempVar1090;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1091 = cyVar537.select(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1092 = cyVar537.getEntireRow(-1, 'asf', true, null, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar577 = tempVar1092;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1093 = cyVar537._recursivelySet(0, 1, 3.5, null, 34343333, 3.5);
                    } catch (err) {}
                    try {
                        tempVar1094 = cyVar537._recursivelySet('asf', 34343333, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1095 = cyVar537.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar578 = tempVar1095;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1096 = cyVar537.getColumn(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar579 = tempVar1096;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1097 = cyVar537.getLastColumn(undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar580 = tempVar1097;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1098 = cyVar537.calculate(34343333, 1, null, 0, 3.5, undefined, 0, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1099 = cyVar537.getOffsetRange(12, 12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar581 = tempVar1099;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1100 = cyVar537.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar582 = tempVar1100;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1101 = cyVar537.getVisibleView(true, 1, 1, undefined, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar583 = tempVar1101;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1102 = cyVar537._KeepReference(0, 34343333, -1, 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1103 = cyVar537._ensureInteger(null, false, 'asf', 0, 1, 3.5);
                    } catch (err) {}
                    try {
                        tempVar1104 = cyVar537._ensureInteger(true, -1, 'asf', 34343333, 3.5, false, false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1105 = cyVar537._handleResult(true, 0, 1, 3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1106 = cyVar537._ValidateArraySize(1, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1107 = cyVar537.getCell(2, 12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar584 = tempVar1107;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1108 = cyVar537.getEntireColumn(null, 1, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar585 = tempVar1108;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1109 = cyVar537.unmerge(null, -1, 34343333, 3.5, 1, 3.5, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1110 = cyVar537._getAdjacentRange(0, 34343333);
                    } catch (err) {}
                    try {
                        tempVar1111 = cyVar537._getAdjacentRange(34343333, true, 3.5, -1, 0, 'asf', 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1112 = cyVar537.getLastRow(false, 0, 1, true, 3.5, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar586 = tempVar1112;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1113 = cyVar536.merge(-1, -1, true, 'asf', true, 'asf', -1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1114 = cyVar536.getEntireRow(-1, -1, true, false, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar587 = tempVar1114;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1115 = cyVar536.getEntireColumn(34343333, undefined, 1, true, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar588 = tempVar1115;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1116 = cyVar536._ValidateArraySize(34343333, true, false, true, 'asf', -1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1117 = cyVar536.getResizedRange(11, 13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar589 = tempVar1117;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1118 = cyVar536.getRowsAbove(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar590 = tempVar1118;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1119 = cyVar536.select(1, false, -1, 0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1120 = cyVar536.untrack(-1, 34343333, undefined, 1, -1, 1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1121 = cyVar536.getColumn(0, 1, null, undefined, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar591 = tempVar1121;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1122 = cyVar536.getLastRow(false, undefined, 'asf', 'asf', 3.5, 1, false, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar592 = tempVar1122;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1123 = cyVar536.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar593 = tempVar1123;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1124 = cyVar536._KeepReference(0, -1, 1, -1, 3.5, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1125 = cyVar536.getCell(5, 10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar594 = tempVar1125;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1126 = cyVar536.getRowsBelow(12);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar595 = tempVar1126;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1127 = cyVar535._KeepReference(1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1128 = cyVar535.getEntireRow(3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar596 = tempVar1128;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1129 = cyVar535.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar597 = tempVar1129;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1130 = cyVar535._getAdjacentRange(34343333, 0, false);
                    } catch (err) {}
                    try {
                        tempVar1131 = cyVar535._getAdjacentRange(34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1132 = cyVar535.getCell(7, 11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar598 = tempVar1132;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1133 = cyVar535.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar599 = tempVar1133;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1134 = cyVar535._ensureInteger(null, 34343333, 1, null);
                    } catch (err) {}
                    try {
                        tempVar1135 = cyVar535._ensureInteger(34343333, true, 1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1136 = cyVar535.untrack(34343333, 34343333, undefined, 0, 3.5, null, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1137 = cyVar535.getRowsBelow(1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar600 = tempVar1137;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1138 = cyVar535.getVisibleView(1, 34343333, null, false, true, 34343333, 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar601 = tempVar1138;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1139 = cyVar535.getLastCell(null, null, 1, 'asf', 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar602 = tempVar1139;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1140 = cyVar535.getColumnsAfter(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar603 = tempVar1140;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1141 = cyVar535.getColumn(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar604 = tempVar1141;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1142 = cyVar535.getRowsAbove(9);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar605 = tempVar1142;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1143 = cyVar535.getColumnsBefore(true, null, false, true, true, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar606 = tempVar1143;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1144 = cyVar535.set(true, null, 1, undefined, 1, null, 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1145 = cyVar535.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar607 = tempVar1145;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1146 = cyVar535.merge(false, 'asf', null, true, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1147 = cyVar535._recursivelySet(3.5, undefined);
                    } catch (err) {}
                    try {
                        tempVar1148 = cyVar535._recursivelySet(undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1149 = cyVar535.unmerge(false, false, 34343333, 34343333, undefined, undefined, 0, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1150 = cyVar535.track(0, true, 1, 34343333, 0, null, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1151 = cyVar535.toJSON('asf', null, true, true, true, 'asf', undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar608 = tempVar1151;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1152 = cyVar535.calculate(null, false, 1, null, undefined, 34343333, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1153 = cyVar535.getLastColumn(false, true, null, false, undefined, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar609 = tempVar1153;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1154 = cyVar535.getUsedRange(null, 1, 3.5, null, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar610 = tempVar1154;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1155 = cyVar535.getResizedRange(3, 13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar611 = tempVar1155;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1156 = cyVar535.getIntersection(1, false, null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar612 = tempVar1156;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1157 = cyVar535.getEntireColumn(0, 0, -1, true, true, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar613 = tempVar1157;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1158 = cyVar535._ValidateArraySize(-1, true, -1, null, -1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1159 = cyVar535.load(1, 0);
                    } catch (err) {}
                    try {
                        tempVar1160 = cyVar535.load(3.5, 'asf', 'asf', null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1161 = cyVar535._handleIdResult(false, false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1162 = cyVar535.getLastRow(34343333, 34343333, 34343333, true, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar614 = tempVar1162;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1163 = cyVar535.getRow(2);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar615 = tempVar1163;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1164 = cyVar534.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar616 = tempVar1164;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1165 = cyVar534._ensureInteger(false, 34343333, 'asf', 0, 3.5);
                    } catch (err) {}
                    try {
                        tempVar1166 = cyVar534._ensureInteger('asf', 1, 34343333, false, 0, null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1167 = cyVar534._recursivelySet('asf', 1, 1, 34343333, -1, 1);
                    } catch (err) {}
                    try {
                        tempVar1168 = cyVar534._recursivelySet('asf', 34343333, true, false, false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1169 = cyVar534._getAdjacentRange(false, undefined, undefined, 34343333);
                    } catch (err) {}
                    try {
                        tempVar1170 = cyVar534._getAdjacentRange('asf', 3.5, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1171 = cyVar534.select('asf', 0, undefined, 'asf', 3.5, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1172 = cyVar534.calculate(false, null, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1173 = cyVar534.getColumnsBefore(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar617 = tempVar1173;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1174 = cyVar534.getLastColumn(null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar618 = tempVar1174;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1175 = cyVar534.getRowsAbove(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar619 = tempVar1175;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1176 = cyVar534.getIntersection('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar620 = tempVar1176;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1177 = cyVar534.unmerge(false, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1178 = cyVar534._KeepReference(null, false, -1, undefined, undefined, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1179 = cyVar534.set(undefined, -1, 1, 0, -1, 34343333, 0);
                    } catch (err) {}
                    try {
                        tempVar1180 = cyVar534.set(null, 3.5, 0, 'asf', 1, false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1181 = cyVar534.track(undefined, 'asf', 'asf', 3.5, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1182 = cyVar534.getLastCell(undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar621 = tempVar1182;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1183 = cyVar534.getColumnsAfter(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar622 = tempVar1183;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1184 = cyVar534.getEntireColumn(true, 34343333, false, false, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar623 = tempVar1184;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1185 = cyVar534.toJSON(-1, true, -1, undefined, 1, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar624 = tempVar1185;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1186 = cyVar534.getEntireRow(34343333, 'asf', 34343333, true, -1, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar625 = tempVar1186;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1187 = cyVar534.getLastRow(true, false, 'asf', 1, 3.5, 34343333, 34343333, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar626 = tempVar1187;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1188 = cyVar534.getIntersectionOrNullObject(-1, false, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar627 = tempVar1188;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1189 = cyVar534.getColumn(undefined, 'asf', 3.5, null, null, false, false, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar628 = tempVar1189;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1190 = cyVar534.getRowsBelow(undefined, 0, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar629 = tempVar1190;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1191 = cyVar534.getRow(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar630 = tempVar1191;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1192 = cyVar534.merge(null, 0, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1193 = cyVar534.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar631 = tempVar1193;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1194 = cyVar534.getUsedRange(null, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar632 = tempVar1194;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1195 = cyVar534._handleIdResult('asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1196 = cyVar534.getResizedRange(0, -1, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar633 = tempVar1196;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1197 = cyVar534.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar634 = tempVar1197;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1198 = cyVar534.getVisibleView(34343333, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar635 = tempVar1198;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1199 = cyVar534.load(true, -1, 3.5, null, 3.5, null);
                    } catch (err) {}
                    try {
                        tempVar1200 = cyVar534.load(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1201 = cyVar534.getOffsetRange(3, 4);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar636 = tempVar1201;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1202 = cyVar534._ValidateArraySize(true, -1, 3.5, undefined, 1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1203 = cyVar534.getCell(3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar637 = tempVar1203;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1204 = cyVar12.getColumnsAfter(7);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar638 = tempVar1204;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1205 = cyVar12.getRowsAbove(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar639 = tempVar1205;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1206 = cyVar12.select(3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1207 = cyVar12._handleIdResult(1, null, 34343333, 0, true, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1208 = cyVar12.calculate(undefined, 'asf', undefined, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1209 = cyVar12.getRowsBelow(11);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar640 = tempVar1209;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1210 = cyVar12.getEntireColumn(1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar641 = tempVar1210;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1211 = cyVar12.insert('Right');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar642 = tempVar1211;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1212 = cyVar12.unmerge(34343333, -1, 3.5, 34343333, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1213 = cyVar12._ensureInteger(0, null, undefined, true, null, 0, 0, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1214 = cyVar12.load(undefined, 3.5, 1, null, undefined, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1215 = cyVar12.getVisibleView(34343333, 3.5, 0, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar643 = tempVar1215;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1216 = cyVar12.getLastCell(false, 3.5, undefined, -1, 1, undefined, 'asf', 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar644 = tempVar1216;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1217 = cyVar12.getColumn(3);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar645 = tempVar1217;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1218 = cyVar12.getOffsetRange(2, 10);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar646 = tempVar1218;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1219 = cyVar12.getLastColumn(null, 'asf', 0, -1, 1, false, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar647 = tempVar1219;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1220 = cyVar12.getUsedRangeOrNullObject(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar648 = tempVar1220;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1221 = cyVar12._handleResult(undefined, 34343333, 1, false, true, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1222 = cyVar12.merge(0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1223 = cyVar12._recursivelySet(null, null);
                    } catch (err) {}
                    try {
                        tempVar1224 = cyVar12._recursivelySet(1, true, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1225 = cyVar12.track(34343333, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1226 = cyVar12.getResizedRange(3.5, null, 3.5, 3.5, 34343333, 34343333, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar649 = tempVar1226;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1227 = cyVar12._ValidateArraySize(true, -1, 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1228 = cyVar12.untrack(-1, -1, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1229 = cyVar12.getUsedRange(null, undefined, 34343333, true, -1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar650 = tempVar1229;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1230 = cyVar12.set(34343333, true, 34343333, 0, 34343333, 3.5, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1231 = cyVar12.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar651 = tempVar1231;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1232 = cyVar12._getAdjacentRange(null, 0, 0, false, null, 3.5, 1);
                    } catch (err) {}
                    try {
                        tempVar1233 = cyVar12._getAdjacentRange('asf', false, 0, null, -1, 34343333, 34343333, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1234 = cyVar12.getRow(13);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar652 = tempVar1234;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1235 = cyVar12.getIntersectionOrNullObject('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar653 = tempVar1235;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1236 = cyVar12.getColumnsBefore(null, false, -1, 34343333, 1, 'asf', 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar654 = tempVar1236;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1237 = cyVar654.getEntireRow(true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar655 = tempVar1237;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1238 = cyVar654.merge(1, 3.5, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1239 = cyVar654.getVisibleView('asf', 0, 0, true, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar656 = tempVar1239;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1240 = cyVar654.getIntersection('A1:B12');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar657 = tempVar1240;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1241 = cyVar654.getLastColumn(false, 34343333, 3.5, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar658 = tempVar1241;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1242 = cyVar654.untrack(1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1243 = cyVar653.getOffsetRange(true, true, 0, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar659 = tempVar1243;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1244 = cyVar653._recursivelySet(34343333, 3.5, undefined, false);
                    } catch (err) {}
                    try {
                        tempVar1245 = cyVar653._recursivelySet(null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1246 = cyVar653._handleIdResult('asf', 34343333, 3.5, 1, 1, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1247 = cyVar653.untrack(null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1248 = cyVar653.getUsedRange(3.5, false, false, false, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar660 = tempVar1248;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1249 = cyVar653.getRow(null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar661 = tempVar1249;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1250 = cyVar652.getOffsetRange(null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar662 = tempVar1250;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1251 = cyVar652.set(-1, true, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1252 = cyVar652.getUsedRange(true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar663 = tempVar1252;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1253 = cyVar652.untrack(true, undefined, null, 0, -1, -1, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1254 = cyVar652.select(true, null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1255 = cyVar652._ensureInteger(true, -1, null, -1, 'asf', 34343333, 34343333, 3.5);
                    } catch (err) {}
                    try {
                        tempVar1256 = cyVar652._ensureInteger(null, false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1257 = cyVar652.getColumnsAfter(0, true);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar664 = tempVar1257;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1258 = cyVar652.getRowsBelow(undefined, false, true, -1, 1, false, 'asf', 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar665 = tempVar1258;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1259 = cyVar652.getLastRow(true, 0, null, 34343333, null, false, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar666 = tempVar1259;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1260 = cyVar652.toJSON(null, false, -1, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar667 = tempVar1260;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1261 = cyVar652.calculate(34343333, -1, 3.5, -1, false, 3.5, undefined);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1262 = cyVar652.getColumn(6);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar668 = tempVar1262;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1263 = cyVar652.insert('Down');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar669 = tempVar1263;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1264 = cyVar652.getResizedRange(11, 5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar670 = tempVar1264;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1265 = cyVar651.load('asf', true, null, 1, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1266 = cyVar651.track('asf', undefined, 1, null, -1, 0);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1267 = cyVar651.getLastCell(34343333, null);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar671 = tempVar1267;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1268 = cyVar651.getUsedRangeOrNullObject(false);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar672 = tempVar1268;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1269 = cyVar651.getBoundingRect('D10:E15');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar673 = tempVar1269;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1270 = cyVar651.untrack(3.5, 1, 3.5, false, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1271 = cyVar650.getRow(5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar674 = tempVar1271;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1272 = cyVar650.getRowsAbove(8);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar675 = tempVar1272;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1273 = cyVar650.getEntireColumn(undefined, -1, 34343333, 3.5, true, 'asf');
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar676 = tempVar1273;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1274 = cyVar650.getIntersection(null, 0, 1, undefined, 3.5, 'asf', -1, 1);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar677 = tempVar1274;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    try {
                        tempVar1275 = cyVar650.load(false, null, 34343333, true, null, null, 3.5);
                    } catch (err) {}
                    try {
                        tempVar1276 = cyVar650.load(1, 1, 3.5);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    try {
                        tempVar1277 = cyVar650.getLastCell(3.5, 34343333, null, 3.5, 1, -1, null, 34343333);
                    } catch (err) {}
                    return context.sync();
                }).catch(function(err) {})
                .then(function() {
                    cyVar678 = tempVar1277;
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    return context.sync();
                }).catch(function(error) {})
                .then(function() {
                    console.log('finished');
                }).catch(function(err) {
                    console.log('finished');
                });
        });
    }
}());