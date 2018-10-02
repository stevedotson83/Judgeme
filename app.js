/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
                console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
            } 
            $('#apply-profile').click(applyProfile);
            $('.profile-buttons .ms-Button').click(function(){
                var selectedTab = $(this).attr('data-target');
                $('.profile-buttons').css('display','none');
                $('.profile-sections .profile-section').fadeOut(0, function(){
                    $(selectedTab).css('display','block');
                });
            })
            $('.back-button').click(function(){
                $('.profile-sections .profile-section').fadeOut(0, function(){
                    $('.profile-buttons').css('display','block');
                });
            })
        });
    };

    function applyProfile() {
        Word.run(function (context) {
            context.document.body.font.set({
                name: "Arial"
            });
            var paras = context.document.body.paragraphs;
            paras.load("items");
            return context.sync().then(function(){
                paras.items.forEach(function(para){
                    para.alignment = "Left";
                })
            });
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
  
})();