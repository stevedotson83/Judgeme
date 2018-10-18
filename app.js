/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {
    var profilesArray = [];
    Office.initialize = function (reason) {
        console.log('ChalBhai');
        console.log(reason);
        $(document).ready(function () {
            if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
                console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
            }; 

            // INITIALIZE
            readProfiles();
            fetchProfiles();
            $('#home-select-profile').focus();
            // UI
            $('.profile-buttons .ms-Button').click(function(){
                console.log('Button Clicked')
                var selectedTab = $(this).attr('data-target');
                $('.profile-buttons').css('display','none');
                $('.profile-sections .profile-section').fadeOut(0, function(){
                    $(selectedTab).css('display','block');
                });
            });
            $('.back-button').click(function(){
                $('.profile-sections .profile-section').fadeOut(0, function(){
                    $('.profile-buttons').css('display','block');
                });
            });
            $('.ms-Pivot-link:nth-of-type(2)').click(function(){
                $('.profile-sections .profile-section').fadeOut(0, function(){
                    $('.profile-buttons').css('display','block');
                });
            });

            // APPLY PROFILE LIST CHANGE
            $('#home-select-profile').change(function (e) { 
                var prefsSettings =  $('#home-select-profile option:selected').attr('value').split("");
                fillChecks('.profile-preferences-section.apply-section', prefsSettings);
            });

            // MODIFY PROFILE LIST CHANGE 
            $('#modify-select-profile').change(function (e) { 
                var prefsSettings =  $('#modify-select-profile option:selected').attr('value').split("");
                fillChecks('.profile-preferences-section.modify-section', prefsSettings);
            });

            // APPLY PROFILE BUTTON CLICK
            $('#apply-profile').click(function(){
                if ($('#home-select-profile option:selected').text() != ""){
                    applyProfile();
                    notifyMessage();
                } else {
                    errorMessage();
                    $('#home-select-profile').focus();
                };
            });

            // CREATE PROFILE BUTTON CLICK
            $('#btn-new-profile').click(function(){
                var name = $('#txt-newprofile-name').val();
                var container = '.profile-preferences-section.add-section';
                addProfile(name, container);
            });

            // MODIFY PROFILE BUTTON CLICK
            $('#btn-edit-profile').click(function(){
                var name = $('#modify-select-profile option:selected').text();
                var container = '.profile-preferences-section.modify-section';
                addProfile(name, container);
            });
            
            // DELETE PROFILE BUTTON CLICK
            $('body').on('click','.btn-deleteprofile', function(){
                var profileToDelete = $(this).parent().parent().find('td').eq(0).text();
                document.cookie = profileToDelete + "=;expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";
                readProfiles();
                fetchProfiles();
            });

        });
    };

    // CHECK / UNCHECK CHECKBOXES ACCORDING TO PROFILE PREFERENCES
    function fillChecks(dataTarget, prefsToCheck){
        for (var index = 0; index < $(dataTarget).find('.check-Pref').length; index++) {
            $(dataTarget).find('.check-Pref').eq(index).prop('checked', parseInt(prefsToCheck[index]));                
        };
    };

    // TO ADD/EDIT A PROFILE IN COOKIES
    function addProfile(profName, checksContainer){
        var profilePrefs = "";
        if (profName != ""){
            for (var a = 0; a < $(checksContainer).find('.check-Pref').length; a++){
                var prefValue = $(checksContainer).find('.check-Pref').eq(a).is(':checked') ? 1:0;
                profilePrefs += prefValue;
            };
            document.cookie = profName + "=" + profilePrefs + ";expires=Sat, 01 Jan 2050 00:00:00 UTC; path=/";
            readProfiles();
            fetchProfiles();
            notifyMessage();
        } else {
            errorMessage();
            
        };
    };

    // TO READ THE SAVED COOKIES AND STORE IN ARRAY
    function readProfiles(){
        profilesArray = [];
        var rawData = decodeURIComponent(document.cookie);
        var profilesRawData =  rawData.split(';');
        for (var i = 2; i < profilesRawData.length ; i++){
            var profile = profilesRawData[i].split('=');
            var profileName = profile[0].substr(1,profile[0].length-1);
            var profilePrefs = profile[1];
            var arrayItem = ({'profilename' : profileName, 'prefs': profilePrefs});
            profilesArray.push(arrayItem);
        };
    };

    // TO SERVE THE SAVED COOKIES INTO THE UI
    function fetchProfiles(){
        $('#home-select-profile').html("");
        $('#modify-select-profile').html("");
        $('.delete-profiles-table tbody').html('');
        for (var i = 0; i < profilesArray.length; i++){
            $('#home-select-profile').append(
                '<option value=' + profilesArray[i].prefs + '>' + profilesArray[i].profilename + '</option>'
            );
            $('#modify-select-profile').append(
                '<option value=' + profilesArray[i].prefs + '>' + profilesArray[i].profilename + '</option>'
            );
            $('.delete-profiles-table tbody').append(
                '<tr><td>'+ profilesArray[i].profilename +'</td><td><button class="ms-Button btn-deleteprofile" title="Delete Profile">Delete</button></td></tr>'
            );
        };   
        refreshUI();
    };
    
    // NOTIFICATION AND ERROR MESSAGES ALERT
    function notifyMessage(){
        $('.notification-message').fadeIn(500, function() {$('.notification-message').fadeOut(5000)});
    };

    function errorMessage(){
        $('.error-message').fadeIn(500, function() {$('.error-message').fadeOut(5000)});
    };

    // REFRESHES THE DROPDOWN LISTS, INPUT TEXT BOX, THE CHECKS 
    function refreshUI(){
        $('#txt-newprofile-name').val('');
        $('#modify-select-profile').val(0);
        $('#home-select-profile').val(profilesArray[0].prefs);
        $(".profile-preferences-section.add-section .check-Pref:checked").prop('checked', false);
        $(".profile-preferences-section.modify-section .check-Pref:checked").prop('checked', false);
        var prefsSettings =  profilesArray[0].prefs.split("");
        fillChecks('.profile-preferences-section.apply-section', prefsSettings);
    }

    // FUNCTION TO APPLY PREFERENCE SETTINGS TO THE DOCUMENT 
    function applyProfile(){
        for (var index = 0; index < $('.profile-preferences-section.apply-section .check-Pref').length; index++) {
            if ($('.profile-preferences-section.apply-section .check-Pref').eq(index).is(':checked') == true){
                var prefID = $('.profile-preferences-section.apply-section .check-Pref').eq(index).attr('target');
                applyProfilePrefrences(prefID);
            };
        };
    };

    // ALL CODE FOR PREFERENCE FUNCTIONALITY GOES INTO THIS FUNCTION
    function applyProfilePrefrences(prefId) {
        Word.run(function (context) {
            if (prefId == "pref1"){        
                context.document.body.font.set({
                    name: "Arial"
                });        
                return context.sync();
            } else if (prefId == "pref2"){
                var paras = context.document.body.paragraphs;
                paras.load("items");
                return context.sync().then(function(){
                    paras.items.forEach(function(para){
                        para.alignment = "Left";
                    });
                });
            } else if (prefId == "pref3"){

            } else if (prefId == "pref4"){

            } else if (prefId == "pref5"){

            };
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            };
        });
    };
})();