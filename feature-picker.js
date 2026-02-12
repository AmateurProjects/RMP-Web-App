/**
 * feature-picker.js
 * AMD module – Feature Picker Modal (overlapping polygon disambiguation)
 *
 * Manages the modal that appears when a user clicks on overlapping polygons,
 * allowing them to preview and select which polygon to use as their AOI.
 *
 * Exports:
 *   init(state, deps)        – wire up DOM, deps = { GraphicsLayer, Graphic }
 *   showFeaturePicker(features, onSelect)
 *   hideFeaturePicker()
 */
define(["app/config-helpers"], function (configHelpers) {
    "use strict";

    var escapeHtml = configHelpers.escapeHtml;

    // ── Module-private state (set by init) ──
    var state = null;   // shared app state  { map }

    // AMD deps
    var GraphicsLayer, Graphic;

    // ── DOM refs (grabbed once in init) ──
    var featurePickerModal, featurePickerList,
        featurePickerCancelBtn, featurePickerConfirmBtn,
        featurePickerContent, featurePickerHeader;

    // ── Internal state ──
    var pickerDragState = { isDragging: false, startX: 0, startY: 0, offsetX: 0, offsetY: 0 };
    var pickerFeatures = [];
    var pickerSelectedIdx = -1;
    var pickerOnSelect = null;
    var pickerHighlightLayer = null;

    // ── Init ──
    function init(appState, deps) {
        state = appState;
        GraphicsLayer = deps.GraphicsLayer;
        Graphic = deps.Graphic;

        // Grab DOM refs
        featurePickerModal   = document.getElementById("featurePickerModal");
        featurePickerList    = document.getElementById("featurePickerList");
        featurePickerCancelBtn  = document.getElementById("featurePickerCancelBtn");
        featurePickerConfirmBtn = document.getElementById("featurePickerConfirmBtn");
        featurePickerContent = document.querySelector(".feature-picker-content");
        featurePickerHeader  = document.querySelector(".feature-picker-header");

        // Wire drag
        if (featurePickerHeader && featurePickerContent) {
            featurePickerHeader.addEventListener("mousedown", function (e) {
                if (e.target.tagName === "BUTTON") return;
                pickerDragState.isDragging = true;
                pickerDragState.startX = e.clientX;
                pickerDragState.startY = e.clientY;
                var rect = featurePickerContent.getBoundingClientRect();
                pickerDragState.offsetX = rect.left;
                pickerDragState.offsetY = rect.top;
                featurePickerContent.style.transition = "none";
            });

            document.addEventListener("mousemove", function (e) {
                if (!pickerDragState.isDragging) return;
                var dx = e.clientX - pickerDragState.startX;
                var dy = e.clientY - pickerDragState.startY;
                var newLeft = pickerDragState.offsetX + dx;
                var newTop  = pickerDragState.offsetY + dy;
                featurePickerContent.style.position = "fixed";
                featurePickerContent.style.left  = newLeft + "px";
                featurePickerContent.style.top   = newTop + "px";
                featurePickerContent.style.right  = "auto";
                featurePickerContent.style.margin = "0";
            });

            document.addEventListener("mouseup", function () {
                pickerDragState.isDragging = false;
                featurePickerContent.style.transition = "";
            });
        }

        // Cancel button
        if (featurePickerCancelBtn) {
            featurePickerCancelBtn.addEventListener("click", hideFeaturePicker);
        }

        // Confirm button
        if (featurePickerConfirmBtn) {
            featurePickerConfirmBtn.addEventListener("click", confirmPickerSelection);
        }

        // Close on click outside
        if (featurePickerModal) {
            featurePickerModal.addEventListener("click", function (e) {
                if (e.target === featurePickerModal) {
                    hideFeaturePicker();
                }
            });
        }
    }

    // ── Display helpers ──

    function getFeaturePickerDisplayName(attrs) {
        var nameFields = [
            "NAME", "Name", "name",
            "ALLOT_NAME", "ALLOTMENT_NAME", "Allotment",
            "PASTURE_NAME", "PASTURE",
            "LEASE_NAME", "LEASE_NUM", "CASE_FILE_N",
            "PLAN_NAME", "UNIT_NAME", "AREA_NAME",
            "LABEL", "TITLE", "DESCRIPTION"
        ];

        for (var i = 0; i < nameFields.length; i++) {
            var field = nameFields[i];
            if (attrs[field] && String(attrs[field]).trim()) {
                return String(attrs[field]).trim();
            }
        }

        // Fallback
        var entries = Object.entries(attrs);
        for (var j = 0; j < entries.length; j++) {
            var key = entries[j][0], val = entries[j][1];
            if (typeof val === "string" && val.trim() &&
                !key.toLowerCase().includes("objectid") &&
                !key.toLowerCase().includes("globalid") &&
                !key.toLowerCase().includes("shape")) {
                return val.trim().substring(0, 60);
            }
        }

        return "Unnamed Feature";
    }

    function getFeaturePickerDetails(attrs) {
        var detailParts = [];
        var skipFields = ["OBJECTID", "GLOBALID", "SHAPE", "SHAPE_LENGTH", "SHAPE_AREA"];

        var count = 0;
        var entries = Object.entries(attrs);
        for (var i = 0; i < entries.length; i++) {
            if (count >= 2) break;
            var key = entries[i][0], val = entries[i][1];
            if (skipFields.some(function (s) { return key.toUpperCase().includes(s); })) continue;
            if (val && String(val).trim() && typeof val !== "object") {
                var displayVal = String(val).trim();
                if (displayVal.length <= 50) {
                    detailParts.push(key + ": " + displayVal);
                    count++;
                }
            }
        }

        return detailParts.join(" \u2022 ");
    }

    // ── Highlight helpers ──

    function highlightPickerFeature(graphic) {
        if (!graphic || !state.map) return;

        clearPickerHighlight();

        pickerHighlightLayer = new GraphicsLayer({ title: "Picker Highlight" });
        state.map.add(pickerHighlightLayer);

        var highlightGraphic = new Graphic({
            geometry: graphic.geometry,
            symbol: {
                type: "simple-fill",
                color: [0, 200, 100, 0.35],
                outline: { color: [0, 150, 50], width: 3 }
            }
        });
        pickerHighlightLayer.add(highlightGraphic);
    }

    function clearPickerHighlight() {
        if (pickerHighlightLayer && state.map) {
            try { state.map.remove(pickerHighlightLayer); } catch (e) {}
            pickerHighlightLayer = null;
        }
    }

    // ── Core picker functions ──

    function showFeaturePicker(features, onSelect) {
        if (!featurePickerModal || !featurePickerList) return;

        // Reset position for new selection
        if (featurePickerContent) {
            featurePickerContent.style.position = "";
            featurePickerContent.style.left = "";
            featurePickerContent.style.top = "";
            featurePickerContent.style.right = "";
            featurePickerContent.style.margin = "";
        }

        // Store state
        pickerFeatures = features;
        pickerSelectedIdx = -1;
        pickerOnSelect = onSelect;

        // Reset confirm button
        if (featurePickerConfirmBtn) {
            featurePickerConfirmBtn.disabled = true;
            featurePickerConfirmBtn.textContent = "Select This Polygon";
        }

        // Build the list of features
        featurePickerList.innerHTML = features.map(function (f, idx) {
            var attrs = f.graphic?.attributes || {};
            var name = getFeaturePickerDisplayName(attrs);
            var details = getFeaturePickerDetails(attrs);

            return '<div class="feature-picker-item" data-idx="' + idx + '">' +
                '<div class="feature-picker-index">' + (idx + 1) + '</div>' +
                '<div class="feature-picker-info">' +
                    '<div class="feature-picker-name">' + escapeHtml(name) + '</div>' +
                    (details ? '<div class="feature-picker-details">' + escapeHtml(details) + '</div>' : '') +
                '</div>' +
            '</div>';
        }).join("");

        // Add click handlers for preview (not immediate selection)
        var items = featurePickerList.querySelectorAll(".feature-picker-item");
        items.forEach(function (item) {
            item.addEventListener("click", function () {
                var idx = parseInt(item.getAttribute("data-idx"), 10);
                selectPickerRow(idx);
            });
        });

        featurePickerModal.classList.remove("hidden");
    }

    function selectPickerRow(idx) {
        if (idx < 0 || idx >= pickerFeatures.length) return;

        pickerSelectedIdx = idx;

        // Update visual selection state
        var items = featurePickerList.querySelectorAll(".feature-picker-item");
        items.forEach(function (item, i) {
            item.classList.toggle("selected", i === idx);
        });

        // Highlight the selected polygon on the map
        highlightPickerFeature(pickerFeatures[idx]?.graphic);

        // Enable confirm button and update text
        if (featurePickerConfirmBtn) {
            featurePickerConfirmBtn.disabled = false;
            var name = getFeaturePickerDisplayName(pickerFeatures[idx]?.graphic?.attributes || {});
            var shortName = name.length > 30 ? name.substring(0, 27) + "..." : name;
            featurePickerConfirmBtn.textContent = 'Select "' + shortName + '"';
        }
    }

    function confirmPickerSelection() {
        if (pickerSelectedIdx < 0 || !pickerFeatures[pickerSelectedIdx]) return;

        var selectedFeature = pickerFeatures[pickerSelectedIdx];
        var callback = pickerOnSelect;
        hideFeaturePicker();

        if (callback) {
            callback(selectedFeature);
        }
    }

    function hideFeaturePicker() {
        if (featurePickerModal) {
            featurePickerModal.classList.add("hidden");
        }
        clearPickerHighlight();
        pickerFeatures = [];
        pickerSelectedIdx = -1;
        pickerOnSelect = null;
    }

    // ── Public API ──
    return {
        init: init,
        showFeaturePicker: showFeaturePicker,
        hideFeaturePicker: hideFeaturePicker
    };
});
