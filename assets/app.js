(function () {
    "use strict";

    function on(element, eventName, handler) {
        if (element.addEventListener) {
            element.addEventListener(eventName, handler, false);
        } else if (element.attachEvent) {
            element.attachEvent("on" + eventName, handler);
        }
    }

    function forEachNode(nodes, callback) {
        var index;
        for (index = 0; index < nodes.length; index += 1) {
            callback(nodes[index], index);
        }
    }

    function initDeleteConfirmation() {
        var forms = document.querySelectorAll("[data-delete-form]");

        forEachNode(forms, function (form) {
            on(form, "submit", function (event) {
                var title = form.getAttribute("data-task-title") || "công việc này";
                var accepted = window.confirm("Xóa ‘" + title + "’? Thao tác này không thể hoàn tác.");

                if (!accepted) {
                    if (event.preventDefault) {
                        event.preventDefault();
                    } else {
                        event.returnValue = false;
                    }
                }
            });
        });
    }

    function initAutoSubmit() {
        var fields = document.querySelectorAll("[data-auto-submit]");

        forEachNode(fields, function (field) {
            on(field, "change", function () {
                if (field.form) {
                    field.form.submit();
                }
            });
        });
    }

    function updateCounter(field) {
        var counter = document.querySelector("[data-counter-for='" + field.id + "']");
        var maximum = field.getAttribute("maxlength") || "";

        if (counter) {
            counter.innerHTML = field.value.length + " / " + maximum;
        }
    }

    function initCounters() {
        var fields = document.querySelectorAll("[data-counted-field]");

        forEachNode(fields, function (field) {
            updateCounter(field);
            on(field, "input", function () {
                updateCounter(field);
            });
        });
    }

    function initSubmitGuard() {
        var forms = document.querySelectorAll("[data-task-form]");

        forEachNode(forms, function (form) {
            on(form, "submit", function (event) {
                var button = form.querySelector("[data-submit-button]");

                if (form.getAttribute("data-submitting") === "true") {
                    if (event.preventDefault) {
                        event.preventDefault();
                    } else {
                        event.returnValue = false;
                    }
                    return false;
                }

                form.setAttribute("data-submitting", "true");
                if (button) {
                    button.disabled = true;
                    button.innerHTML = "Đang lưu…";
                }

                return true;
            });
        });
    }

    function init() {
        initDeleteConfirmation();
        initAutoSubmit();
        initCounters();
        initSubmitGuard();
    }

    if (document.readyState === "loading") {
        on(document, "DOMContentLoaded", init);
    } else {
        init();
    }
}());
