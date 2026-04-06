"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_http_1 = require("@microsoft/sp-http");
var sp_property_pane_1 = require("@microsoft/sp-property-pane");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var KIHubPromptCardsWebPart_module_scss_1 = tslib_1.__importDefault(require("./KIHubPromptCardsWebPart.module.scss"));
var KIHubPromptCardsWebPart = /** @class */ (function (_super) {
    tslib_1.__extends(KIHubPromptCardsWebPart, _super);
    function KIHubPromptCardsWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    KIHubPromptCardsWebPart.prototype.render = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var items, error_1;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.domElement.innerHTML = "\n      <section class=\"".concat(KIHubPromptCardsWebPart_module_scss_1.default.promptCards, "\">\n        <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.headerBlock, "\">\n          <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.kicker, "\">Copilot Prompt Library</div>\n          <h2 class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.mainTitle, "\">Prompt Cards</h2>\n          <p class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.mainSubtitle, "\">\n            Start with a polished prompt, copy it quickly, and launch Copilot in one click.\n          </p>\n        </div>\n\n        <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.loadingState, "\">\n          Loading prompts...\n        </div>\n      </section>\n    ");
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._getPromptItems()];
                    case 2:
                        items = _a.sent();
                        this._renderCards(items);
                        this._wireUpEvents();
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        console.error(error_1);
                        this.domElement.innerHTML = "\n        <section class=\"".concat(KIHubPromptCardsWebPart_module_scss_1.default.promptCards, "\">\n          <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.errorState, "\">\n            Unable to load the Copilot Prompt Library right now.\n          </div>\n        </section>\n      ");
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    KIHubPromptCardsWebPart.prototype._getPromptItems = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var listName, filterType, filterValue, filterQuery, endpoint, response, json;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        listName = this.properties.listName || 'Copilot Prompt Library';
                        filterType = this.properties.filterType || 'None';
                        filterValue = this.properties.filterValue || '';
                        filterQuery = '';
                        if (filterType === 'Featured') {
                            filterQuery = "&$filter=Featured eq 1";
                        }
                        else if (filterType === 'ProgramArea' && filterValue) {
                            filterQuery = "&$filter=ProgramAreas eq '".concat(filterValue.replace(/'/g, "''"), "'");
                        }
                        endpoint = "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists/getbytitle('").concat(encodeURIComponent(listName), "')/items") +
                            "?$select=Id,Title,field_1,field_2,ProgramAreas,Featured,PromptLink" +
                            "".concat(filterQuery) +
                            "&$orderby=Featured desc,Id asc";
                        return [4 /*yield*/, this.context.spHttpClient.get(endpoint, sp_http_1.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Error loading prompt items: ".concat(response.status, " ").concat(response.statusText));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        return [2 /*return*/, json.value || []];
                }
            });
        });
    };
    KIHubPromptCardsWebPart.prototype._escapeHtml = function (value) {
        if (!value) {
            return '';
        }
        return value
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    };
    KIHubPromptCardsWebPart.prototype._truncate = function (value, maxLength) {
        if (!value) {
            return '';
        }
        return value.length > maxLength ? "".concat(value.substring(0, maxLength), "...") : value;
    };
    KIHubPromptCardsWebPart.prototype._resolvePromptLink = function (item) {
        var fallback = this.properties.copilotUrl || 'https://m365.cloud.microsoft/chat';
        var raw = item.PromptLink;
        if (!raw) {
            return fallback;
        }
        if (typeof raw === 'string') {
            return raw.startsWith('http') ? raw : fallback;
        }
        if (raw.Url && raw.Url.startsWith('http')) {
            return raw.Url;
        }
        return fallback;
    };
    KIHubPromptCardsWebPart.prototype._renderCards = function (items) {
        var _this = this;
        if (!items.length) {
            this.domElement.innerHTML = "\n        <section class=\"".concat(KIHubPromptCardsWebPart_module_scss_1.default.promptCards, "\">\n          <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.emptyState, "\">\n            No prompt items matched this page filter.\n          </div>\n        </section>\n      ");
            return;
        }
        var cardsHtml = items.map(function (item) {
            var title = _this._escapeHtml(item.Title || 'Untitled Prompt');
            var category = _this._escapeHtml(item.ProgramAreas || 'General');
            var beginnerPrompt = _this._escapeHtml(item.field_1 || '');
            var advancedPrompt = _this._escapeHtml(item.field_2 || '');
            var featured = !!item.Featured;
            var promptLink = _this._escapeHtml(_this._resolvePromptLink(item));
            var initialPrompt = beginnerPrompt || advancedPrompt || 'No prompt text available.';
            var previewPrompt = _this._truncate(initialPrompt, 320);
            return "\n        <article class=\"".concat(KIHubPromptCardsWebPart_module_scss_1.default.card, "\">\n          <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.cardHeader, "\">\n            <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.badgeRow, "\">\n              <span class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.categoryBadge, "\">").concat(category, "</span>\n              ").concat(featured ? "<span class=\"".concat(KIHubPromptCardsWebPart_module_scss_1.default.featuredBadge, "\">Featured</span>") : '', "\n            </div>\n\n            <h3 class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.cardTitle, "\">").concat(title, "</h3>\n          </div>\n\n          <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.modeRow, "\">\n            <button\n              type=\"button\"\n              class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.modeButton, " ").concat(KIHubPromptCardsWebPart_module_scss_1.default.modeButtonActive, "\"\n              data-role=\"mode\"\n              data-mode=\"beginner\"\n              data-card-id=\"").concat(item.Id, "\">\n              Beginner\n            </button>\n\n            <button\n              type=\"button\"\n              class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.modeButton, "\"\n              data-role=\"mode\"\n              data-mode=\"advanced\"\n              data-card-id=\"").concat(item.Id, "\">\n              Advanced\n            </button>\n          </div>\n\n          <div\n            class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.promptBody, "\"\n            id=\"prompt-body-").concat(item.Id, "\"\n            data-beginner=\"").concat(beginnerPrompt, "\"\n            data-advanced=\"").concat(advancedPrompt, "\"\n            data-current=\"").concat(initialPrompt, "\">\n            ").concat(previewPrompt, "\n          </div>\n\n          <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.actionRow, "\">\n            <button\n              type=\"button\"\n              class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.secondaryButton, "\"\n              data-role=\"copy\"\n              data-card-id=\"").concat(item.Id, "\">\n              Copy Prompt\n            </button>\n\n            <button\n              type=\"button\"\n              class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.primaryButton, "\"\n              data-role=\"copilot\"\n              data-card-id=\"").concat(item.Id, "\"\n              data-link=\"").concat(promptLink, "\">\n              Use in Copilot\n            </button>\n          </div>\n        </article>\n      ");
        }).join('');
        this.domElement.innerHTML = "\n      <section class=\"".concat(KIHubPromptCardsWebPart_module_scss_1.default.promptCards, "\">\n        <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.headerBlock, "\">\n          <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.kicker, "\">Copilot Prompt Library</div>\n          <h2 class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.mainTitle, "\">Prompt Cards</h2>\n          <p class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.mainSubtitle, "\">\n            Start with a polished prompt, copy it quickly, and launch Copilot in one click.\n          </p>\n        </div>\n\n        <div class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.grid, "\">\n          ").concat(cardsHtml, "\n        </div>\n\n        <div id=\"kihub-toast\" class=\"").concat(KIHubPromptCardsWebPart_module_scss_1.default.toast, "\" aria-live=\"polite\"></div>\n      </section>\n    ");
    };
    KIHubPromptCardsWebPart.prototype._wireUpEvents = function () {
        var _this = this;
        var modeButtons = this.domElement.querySelectorAll('button[data-role="mode"]');
        modeButtons.forEach(function (button) {
            button.addEventListener('click', function () {
                var cardId = button.getAttribute('data-card-id') || '';
                var mode = button.getAttribute('data-mode') || 'beginner';
                var promptBody = _this.domElement.querySelector("#prompt-body-".concat(cardId));
                if (!promptBody) {
                    return;
                }
                var beginnerPrompt = promptBody.getAttribute('data-beginner') || '';
                var advancedPrompt = promptBody.getAttribute('data-advanced') || '';
                var selectedPrompt = '';
                if (mode === 'advanced') {
                    selectedPrompt = advancedPrompt || beginnerPrompt || 'No prompt text available.';
                }
                else {
                    selectedPrompt = beginnerPrompt || advancedPrompt || 'No prompt text available.';
                }
                promptBody.setAttribute('data-current', selectedPrompt);
                promptBody.textContent = _this._truncate(selectedPrompt, 320);
                var siblingButtons = _this.domElement.querySelectorAll("button[data-role=\"mode\"][data-card-id=\"".concat(cardId, "\"]"));
                siblingButtons.forEach(function (sibling) {
                    sibling.classList.remove(KIHubPromptCardsWebPart_module_scss_1.default.modeButtonActive);
                });
                button.classList.add(KIHubPromptCardsWebPart_module_scss_1.default.modeButtonActive);
            });
        });
        var copyButtons = this.domElement.querySelectorAll('button[data-role="copy"]');
        copyButtons.forEach(function (button) {
            button.addEventListener('click', function () { return tslib_1.__awaiter(_this, void 0, void 0, function () {
                var cardId, promptBody, prompt;
                return tslib_1.__generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            cardId = button.getAttribute('data-card-id') || '';
                            promptBody = this.domElement.querySelector("#prompt-body-".concat(cardId));
                            prompt = (promptBody === null || promptBody === void 0 ? void 0 : promptBody.getAttribute('data-current')) || '';
                            return [4 /*yield*/, this._copyPrompt(prompt)];
                        case 1:
                            _a.sent();
                            return [2 /*return*/];
                    }
                });
            }); });
        });
        var copilotButtons = this.domElement.querySelectorAll('button[data-role="copilot"]');
        copilotButtons.forEach(function (button) {
            button.addEventListener('click', function () { return tslib_1.__awaiter(_this, void 0, void 0, function () {
                var cardId, promptBody, prompt, targetUrl;
                return tslib_1.__generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            cardId = button.getAttribute('data-card-id') || '';
                            promptBody = this.domElement.querySelector("#prompt-body-".concat(cardId));
                            prompt = (promptBody === null || promptBody === void 0 ? void 0 : promptBody.getAttribute('data-current')) || '';
                            targetUrl = button.getAttribute('data-link') ||
                                this.properties.copilotUrl ||
                                'https://m365.cloud.microsoft/chat';
                            return [4 /*yield*/, this._copyPrompt(prompt)];
                        case 1:
                            _a.sent();
                            this._showToast('Prompt copied. Opening Copilot...');
                            window.open(targetUrl, '_blank');
                            return [2 /*return*/];
                    }
                });
            }); });
        });
    };
    KIHubPromptCardsWebPart.prototype._copyPrompt = function (prompt) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var error_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, navigator.clipboard.writeText(prompt)];
                    case 1:
                        _a.sent();
                        this._showToast('Prompt copied.');
                        return [3 /*break*/, 3];
                    case 2:
                        error_2 = _a.sent();
                        console.error('Clipboard copy failed.', error_2);
                        this._showToast('Copy failed. Please copy manually.');
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    KIHubPromptCardsWebPart.prototype._showToast = function (message) {
        var toast = this.domElement.querySelector('#kihub-toast');
        if (!toast) {
            return;
        }
        toast.textContent = message;
        toast.classList.add(KIHubPromptCardsWebPart_module_scss_1.default.toastVisible);
        window.setTimeout(function () {
            toast.classList.remove(KIHubPromptCardsWebPart_module_scss_1.default.toastVisible);
        }, 2200);
    };
    Object.defineProperty(KIHubPromptCardsWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    KIHubPromptCardsWebPart.prototype.getPropertyPaneConfiguration = function () {
        var filterOptions = [
            { key: 'None', text: 'None' },
            { key: 'Featured', text: 'Featured only' },
            { key: 'ProgramArea', text: 'Program Area' }
        ];
        return {
            pages: [
                {
                    header: {
                        description: 'Prompt Cards Settings'
                    },
                    groups: [
                        {
                            groupName: 'Data',
                            groupFields: [
                                (0, sp_property_pane_1.PropertyPaneTextField)('listName', {
                                    label: 'List Name'
                                }),
                                (0, sp_property_pane_1.PropertyPaneTextField)('copilotUrl', {
                                    label: 'Default Copilot URL'
                                }),
                                (0, sp_property_pane_1.PropertyPaneDropdown)('filterType', {
                                    label: 'Filter Type',
                                    options: filterOptions
                                }),
                                (0, sp_property_pane_1.PropertyPaneTextField)('filterValue', {
                                    label: 'Filter Value'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return KIHubPromptCardsWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = KIHubPromptCardsWebPart;
//# sourceMappingURL=KIHubPromptCardsWebPart.js.map