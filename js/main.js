(function(window, $) {
    'use strict';

    // Cache document for fast access.
    var document = window.document;


    $('a.toggle-menu').click(function() {
        if ($('ul.menu').find('li:visible').length == 1) {
            $('ul.menu').find('li').show();
        } else {
            $('ul.menu').find('li').not(":eq(0)").hide();
        }
    });


    var owl = $("#owl-demo");

    owl.owlCarousel({
        items: 6
    });

    // Custom Navigation Events
    $(".next").click(function() {
        owl.trigger('owl.next');
    })
    $(".prev").click(function() {
        owl.trigger('owl.prev');
    })
})(window, jQuery);

jQuery(function() {
    // nav menu event
    jQuery(".toggle-menu").click(function() {
        var state = jQuery(this).attr("state");
        if (state == 0) {
            jQuery(this).attr("state", "1");
            jQuery(".c_nav_menu").show();
        } else {
            jQuery(this).attr("state", "0");
            jQuery(".c_nav_menu").hide();
        }
    });

    // window resize
    jQuery(window).resize(function() {
        var width = jQuery(window).innerWidth();
        if (width > 999) {
            jQuery(".c_nav_menu").show();
        } else {
            jQuery(".c_nav_menu").hide();
        }
    });
});

$(document).ready(function() {
    var itaImgLink = "/images/flg/ita.gif";
    var engImgLink = "/images/flg/eng.gif";
    var deuImgLink = "/images/flg/deu.gif";
    var fraImgLink = "/images/flg/fra.gif";

    var imgBtnSel = $('#imgBtnSel');
    var imgBtnIta = $('#imgBtnIta');
    var imgBtnEng = $('#imgBtnEng');
    var imgBtnDeu = $('#imgBtnDeu');
    var imgBtnFra = $('#imgBtnFra');

    var imgNavSel = $('#imgNavSel');
    var imgNavIta = $('#imgNavIta');
    var imgNavEng = $('#imgNavEng');
    var imgNavDeu = $('#imgNavDeu');
    var imgNavFra = $('#imgNavFra');

    var spanNavSel = $('#lanNavSel');
    var spanBtnSel = $('#lanBtnSel');

    imgBtnSel.attr("src", itaImgLink);
    imgBtnIta.attr("src", itaImgLink);
    imgBtnEng.attr("src", engImgLink);
    imgBtnDeu.attr("src", deuImgLink);
    imgBtnFra.attr("src", fraImgLink);

    imgNavSel.attr("src", itaImgLink);
    imgNavIta.attr("src", itaImgLink);
    imgNavEng.attr("src", engImgLink);
    imgNavDeu.attr("src", deuImgLink);
    imgNavFra.attr("src", fraImgLink);

    $(".language").on("click", function(event) {
        var currentId = $(this).attr('id');

        if (currentId == "navIta") {
            imgNavSel.attr("src", itaImgLink);
            spanNavSel.text("ITA");
        } else if (currentId == "navEng") {
            imgNavSel.attr("src", engImgLink);
            spanNavSel.text("ENG");
        } else if (currentId == "navDeu") {
            imgNavSel.attr("src", deuImgLink);
            spanNavSel.text("DEU");
        } else if (currentId == "navFra") {
            imgNavSel.attr("src", fraImgLink);
            spanNavSel.text("FRA");
        }

        if (currentId == "btnIta") {
            imgBtnSel.attr("src", itaImgLink);
            spanBtnSel.text("ITA");
        } else if (currentId == "btnEng") {
            imgBtnSel.attr("src", engImgLink);
            spanBtnSel.text("ENG");
        } else if (currentId == "btnDeu") {
            imgBtnSel.attr("src", deuImgLink);
            spanBtnSel.text("DEU");
        } else if (currentId == "btnFra") {
            imgBtnSel.attr("src", fraImgLink);
            spanBtnSel.text("FRA");
        }

    });
});