/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};

/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {

/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId])
/******/ 			return installedModules[moduleId].exports;

/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			exports: {},
/******/ 			id: moduleId,
/******/ 			loaded: false
/******/ 		};

/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);

/******/ 		// Flag the module as loaded
/******/ 		module.loaded = true;

/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}


/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;

/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;

/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";

/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ function(module, exports) {

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
	            $('.profile-buttons .ms-Button').click(function () {
	                var selectedTab = $(this).attr('data-target');
	                $('.profile-buttons').css('display', 'none');
	                $('.profile-sections .profile-section').fadeOut(0, function () {
	                    $(selectedTab).css('display', 'block');
	                });
	            });
	            $('.back-button').click(function () {
	                $('.profile-sections .profile-section').fadeOut(0, function () {
	                    $('.profile-buttons').css('display', 'block');
	                });
	            });
	        });
	    };

	    function applyProfile() {
	        Word.run(function (context) {
	            context.document.body.font.set({
	                name: "Arial"
	            });
	            var paras = context.document.body.paragraphs;
	            paras.load("items");
	            return context.sync().then(function () {
	                paras.items.forEach(function (para) {
	                    para.alignment = "Left";
	                });
	            });
	        }).catch(function (error) {
	            console.log("Error: " + error);
	            if (error instanceof OfficeExtension.Error) {
	                console.log("Debug info: " + JSON.stringify(error.debugInfo));
	            }
	        });
	    }
	})();

/***/ }
/******/ ]);