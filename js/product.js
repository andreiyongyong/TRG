$(function() {
    var thumbList = $('.product-thumb');
    for (var i = 0; i < thumbList.length; i++) {
        var obj = $(thumbList[i]);
        if (i <= 3) {
            obj.show();
        } else {
            obj.hide();
        }
    }
});

$(document).ready(function() {
    $('.upper-triangle').click(function() {
        var thumbList = $('.product-thumb');
        var first = $('.product-thumb:visible').first().attr("no") - 2;
        if (first != 0) {
            for (var i = 0; i < thumbList.length; i++) {
                var obj = $(thumbList[i]);
                if (i >= first - 1 && i <= (first + 2)) {
                    obj.show();
                } else {
                    obj.hide();
                }
            }
        }
    });

    $('.bottom-triangle').click(function() {
        var thumbList = $('.product-thumb');
        var first = $('.product-thumb:visible').first().attr("no") - 2;
        if (first < thumbList.length - 4) {
            for (var i = 0; i < thumbList.length; i++) {
                var obj = $(thumbList[i]);
                if (i > first && i <= (first + 4)) {
                    obj.show();
                } else {
                    obj.hide();
                }
            }
        }
    })


})