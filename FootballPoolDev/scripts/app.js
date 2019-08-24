/**
 * Created with JetBrains WebStorm.
 * User: adidaska
 * Date: 11/13/13
 * Time: 2:42 AM
 * To change this template use File | Settings | File Templates.
 */
$(document).ready(function(){
    //
    $('.spreadContainer').each(function(){
        var spreadValue = $(this).text().trim().substr(0,1);
        var spreadValText = $(this).text().trim();
        if (spreadValue == '-'){
            // negative -> left
            $(this).addClass('spreadLeft');
        } else if(spreadValue == '+') {
            // positive -> right
            $(this).addClass('spreadRight');
            $(this).find('.spread').text(spreadValText.replace('+', '-'));
        } else {
            // even
            $(this).addClass('spreadEven');
        }
    });

    $('.spreadContainer').each(function(){
        var temp = $(this).find('.spread');
        temp.text(temp.text().trim())
    });

    $("input[value='Update']").click(function(){
        $("#processing").show()
    })


})