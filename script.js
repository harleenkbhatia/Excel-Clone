const PS = new PerfectScrollbar('#cells', {
    wheelSpeed: 12,
    wheelPropagation: true,
});
function findRowCol(ele) {
    let idArray = $(ele).attr("id").split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [rowId, colId];
}

for (let i = 1; i <= 100; i++) {
    let str = "";
    let n = i;

    while (n > 0) {
        let rem = n % 26;
        if (rem == 0) {
            str = 'Z' + str;
            n = Math.floor((n / 26)) - 1;
        }
        else {
            str = String.fromCharCode((rem - 1) + 65) + str;
            n = Math.floor((n / 26));
        }
    }
    $("#columns").append(`<div class = "column-name">${str}</div>`);
    $("#rows").append(`<div class = "row-name">${i}</div>`);
}
for (let i = 1; i <= 100; i++) {
    let row = $(`<div class = "cell-row"></div>`);
    for (let j = 1; j <= 100; j++) {
        row.append(`<div id ="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`)
    }
    $("#cells").append(row);
}
$("#cells").scroll(function () {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
});
$(".input-cell").dblclick(function () {
    $(this).attr("contenteditable", "true");
    $(this).focus();
});
$(".input-cell").blur(function () {
    $(this).attr("contenteditable", "false");
});

function getTopBottomLeftRight(rowId, colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);
    return [topCell, bottomCell, leftCell, rightCell];
}

$(".input-cell").click(function (e) {
    let [rowId, colId] = findRowCol(this); //mapping
    let [topCell, bottomCell, leftCell, rightCell] = getTopBottomLeftRight(rowId, colId);

    if ($(this).hasClass("selected") && e.ctrlKey) {
        deselectCell(this, e, topCell, bottomCell, leftCell, rightCell)
    }
    else {
        selectCell(this, e, topCell, bottomCell, leftCell, rightCell);
    }

})

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell, mouseSelection) {
    if (e.ctrlKey || mouseSelection) {

        //top selected or not
        let topSelected;
        if (topCell) {
            topSelected = topCell.hasClass("selected");
        }

        //bottom selected or not
        let bottomSelected = bottomCell.hasClass("selected");

        //left selected or not
        let leftSelected;
        if (leftCell) {
            leftSelected = leftCell.hasClass("selected");
        }

        //right selected or not
        let rightSelected = rightCell.hasClass("selected");

        if (topSelected) {
            topCell.addClass("bottom-selected");
            $(ele).addClass("top-selected");
        }

        if (leftSelected) {
            leftCell.addClass("right-selected");
            $(ele).addClass("left-selected");
        }

        if (rightSelected) {
            rightCell.addClass("left-selected");
            $(ele).addClass("right-selected");
        }

        if (bottomSelected) {
            bottomCell.addClass("top-selected");
            $(ele).addClass("bottom-selected");
        }

    } else {
        $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    }

    $(ele).addClass("selected");
}
function deselectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {
    if ($(ele).attr("contenteditable") == "false") {
        if ($(ele).hasClass("top-selected")) {
            topCell.removeClass("bottom-selected");
        }
        if ($(ele).hasClass("left-selected")) {
            leftCell.removeClass("right-selected");
        }
        if ($(ele).hasClass("right-selected")) {
            rightCell.removeClass("left-selected");
        }
        if ($(ele).hasClass("bottom-selected")) {
            bottomCell.removeClass("top-selected");
        }
        $(ele).removeClass("selected top-selected bottom-selected left-selected right-selected");
    }

}
let mouseMoved = false;
let startCellStored = false;
let startCell;
let endCell;
$(".input-cell").mousemove(function (event) {
    event.preventDefault();
    if (event.buttons == 1 && !event.ctrlKey) {
        $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
        mouseMoved = true;
        if (!startCellStored) {
            let [rowId, colId] = findRowCol(event.target);
            startCell = { rowId: rowId, colId: colId };
            startCellStored = true;
        } else {
            let [rowId, colId] = findRowCol(event.target);
            endCell = { rowId: rowId, colId: colId };
            selectAllInRange(startCell, endCell);
        }
    } else if (event.buttons == 0 && mouseMoved) {
        startCellStored = false;
        mouseMoved = false;
    }
})
function selectAllInRange(start, end) {
    for (let i = (start.rowId < end.rowId ? start.rowId : end.rowId); i <= (start.rowId < end.rowId ? end.rowId : start.rowId); i++) {
        for (let j = (start.colId < end.colId ? start.colId : end.colId); j <= (start.colId < end.colId ? end.colId : start.colId); j++) {
            let [topCell, bottomCell, leftCell, rightCell] = getTopBottomLeftRight(i, j);
            selectCell($(`#row-${i}-col-${j}`)[0], {}, topCell, bottomCell, leftCell, rightCell, true);
        }
    }
}
$("#bold").click(function(e){
    if($(this).hasClass("selected")){
        $(this).removeClass("selected");
        $(".input-cell.selected").each(function(index, ele){ //same as for each loop
            $(ele).html(`${$(ele).text()}`);
        })
    }else{
        $(this).addClass("selected");
        //we need to add <b> hdhd </b> this tag to the text for making it  bold
        $(".input-cell.selected").each(function(index, ele){ //same as for each loop
            $(ele).html(`<b>${$(ele).text()}</b>`);
        })
    }
})
$("#italic").click(function(e){
    if($(this).hasClass("selected")){
        $(this).removeClass("selected");
        $(".input-cell.selected").each(function(index, ele){ //same as for each loop
            $(ele).html(`${$(ele).text()}`);
        })
    }else{
        $(this).addClass("selected");
        //we need to add <b> hdhd </b> this tag to the text for making it  bold
        $(".input-cell.selected").each(function(index, ele){ //same as for each loop
            $(ele).html(`<i>${$(ele).text()}</i>`);
        })
    }
    
})