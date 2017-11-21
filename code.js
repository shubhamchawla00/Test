'use strict';

function AdaptiveCardMobileRender(targetDom) {
    this.targetDom = targetDom || "#content";
 }

AdaptiveCardMobileRender.prototype.HttpAction = function () {
    AdaptiveCards.HttpAction.call(this);
    this._hideCardOnInvoke = false;
    this._isAutoInvokeAction = false;
    this._autoInvokeOptions = null;

    Object.defineProperties(this, {
        "hideCardOnInvoke": {
            get: function () {
                return this._hideCardOnInvoke;
            },
            enumerable: true,
            configurable: true
        },
        "isAutoInvokeAction": {
            get: function () {
                return this._isAutoInvokeAction;
            },
            enumerable: true,
            configurable: true
        },
        "autoInvokeOptions": {
            get: function () {
                return this._autoInvokeOptions;
            },
            enumerable: true,
            configurable: true
        }
    });
};

AdaptiveCardMobileRender.prototype.init = function () {
    AdaptiveCards.AdaptiveCard.onExecuteAction = onExecuteAction;

    AdaptiveCards.AdaptiveCard.actionTypeRegistry.unregisterType("Action.Submit"); // Action.Submit is not supported in Mobile

    // -------------------------- Customize http action for Mobile ---------------------------
    this.HttpAction.prototype = Object.create(AdaptiveCards.HttpAction.prototype);
    this.HttpAction.prototype.parse = function (json) {
        AdaptiveCards.HttpAction.prototype.parse.call(this, json);
        this._hideCardOnInvoke = json["hideCardOnInvoke"] || false;
        this._isAutoInvokeAction = json["isAutoInvokeAction"] || false;
        if (this._isAutoInvokeAction) {
            this._autoInvokeOptions = json["autoInvokeOptions"];
        }
    };

    this.HttpAction.prototype.prepare = function (inputs) {
        if (this._originalData) {
            this._processedData = JSON.parse(JSON.stringify(this._originalData));
        }
        else {
            this._processedData = {};
        }
        for (var i = 0; i < inputs.length; i++) {
            var inputValue = inputs[i].value;
            if (inputValue != null) {
                this._processedData[inputs[i].id] = inputs[i].value;
            }
        }
        this._isPrepared = true;
    };

    Object.defineProperty(this.HttpAction.prototype, "data", {
        get: function () {
            return this._isPrepared ? this._processedData : this._originalData;
        },
        set: function (value) {
            this._originalData = value;
            this._isPrepared = false;
        },
        enumerable: true,
        configurable: true
    });

    AdaptiveCards.AdaptiveCard.actionTypeRegistry.registerType("Action.Http", function () {
        return new this.HttpAction();
    }.bind(this));
};

AdaptiveCardMobileRender.onExecuteAction = null;

AdaptiveCardMobileRender.prototype.registerActionExecuteCallback = function (callback) {
    AdaptiveCardMobileRender.onExecuteAction = callback;
};

AdaptiveCardMobileRender.prototype.render = function (card) {
    var messageCard = new MessageCard();
    messageCard.parse(card);
    var renderedCard = messageCard.render();
    var parent = document.querySelector(this.targetDom);
    parent.innerHTML = '';
    parent.appendChild(renderedCard);
};

AdaptiveCardMobileRender.prototype.onActionExecuted = function (action, request) {
    var successText = "The action completed successfully.";
    var failureText = "An error occurred. Please try again later.";
    var displayText = successText;
    if (request.status === 200) {
        var responseText = request.responseText;
        try {
            var response = JSON.parse(responseText);
            if (response["status"] !== "Completed") {
                displayText = failureText;
            }
        } catch (e) {
            displayText = failureText;
        }
    } else {
        displayText = failureText;
    }

    action.setStatus(buildStatusCard(displayText, "normal", "large"));
};

function buildStatusCard(text, weight, size) {
    return {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "TextBlock",
                "text": text,
                "weight": weight,
                "size": size
            }
        ]
    };
};

var temp = null;
function onExecuteAction(action) {
    var message = "Action executed\n";
    message += "    Title: " + action.title + "\n";
    if (action instanceof AdaptiveCards.ShowCardAction){
        temp = action
        showPopupCard(action);
    }
    else if (action instanceof AdaptiveCards.OpenUrlAction) {
        message += "    Type: OpenUrl\n";
        message += "    Url: " + action.url + "\n";
    }
    else if (action instanceof AdaptiveCards.HttpAction) {
        message += "    Type: Submit";
        message += "    Data: " + JSON.stringify(action.data);
        if (AdaptiveCardMobileRender.onExecuteAction != null){
            AdaptiveCardMobileRender.onExecuteAction;
        }

        temp.setStatus(
        {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "Working on it...",
                    "weight": "normal",
                    "size": "small"
                }
            ]
        });
    }
    else {
        message += "    Type: <unknown>";
    }

    alert(message);
}

function showPopupCard(action) {
    var width = 350;
    var height = 250;
    // We are running in the browser so we need to center the new window ourselves
    var left = window.screenLeft ? window.screenLeft : window.screenX;
    var top = window.screenTop ? window.screenTop : window.screenY;
    left += (window.innerWidth / 2) - (width / 2);
    top += (window.innerHeight / 2) - (height / 2);
    // Open a child window with a desired set of standard browser features
    var popupWindow = window.open("", '_blank', 'toolbar=no, location=yes, status=no, menubar=no, top=' + top + ', left=' + left + ', width=' + width + ', height=' + height);
    if (!popupWindow) {
        // If we failed to open the window fail the authentication flow
        throw new Error("Failed to open popup");
    };

    //TODO: Change this as required
    popupWindow.document.head.innerHTML+= '<link rel="stylesheet" type="text/css" href="E:/Work/AdaptiveCards/source/nodejs/adaptivecards-visualizer/css/app.css">';
    popupWindow.document.head.innerHTML+= '<link rel="stylesheet" type="text/css" href="E:/Work/AdaptiveCards/source/nodejs/adaptivecards-visualizer/css/teams.css">';

    var overlayElement = popupWindow.document.createElement("div");
    overlayElement.id = "popupOverlay";
    overlayElement.className = "popupOverlay";
    overlayElement.tabIndex = 0;
    overlayElement.style.width = "auto"; // popupWindow.document.documentElement.scrollWidth + "px";
    overlayElement.style.height = popupWindow.document.documentElement.scrollHeight + "px";
    overlayElement.onclick = function (e) {
        document.body.removeChild(overlayElement);
    };
    var cardContainer = popupWindow.document.createElement("div");
    cardContainer.className = "popupCardContainer";
    cardContainer.onclick = function (e) { e.stopPropagation(); };
    cardContainer.appendChild(action.card.render());
    overlayElement.appendChild(cardContainer);
    popupWindow.document.body.appendChild(overlayElement);
}

function MessageCard() {
    this.style = "default";
}

function HostContainer() {
    this.allowCardTitle = true;
    this.allowFacts = true;
    this.allowHeroImage = true;
    this.allowImages = true;
    this.allowActionCard = false;
}

MessageCard.prototype.parse = function (json) {
    this.hostContainer = new HostContainer();

    this.defaultCardConfig = {
        "supportsInteractivity": true,
        "fontFamily": "Segoe UI",
        "fontSizes": {
            "small": 12,
            "default": 14,
            "medium": 17,
            "large": 21,
            "extraLarge": 26
        },
        "fontWeights": {
            "lighter": 200,
            "default": 400,
            "bolder": 600
        },
        "imageSizes": {
            "small": 40,
            "medium": 80,
            "large": 160
        },
        "containerStyles": {
            "default": {
                "fontColors": {
                    "default": {
                        "normal": "#333333",
                        "subtle": "#EE333333"
                    },
                    "accent": {
                        "normal": "#2E89FC",
                        "subtle": "#882E89FC"
                    },
                    "good": {
                        "normal": "#54a254",
                        "subtle": "#DD54a254"
                    },
                    "warning": {
                        "normal": "#e69500",
                        "subtle": "#DDe69500"
                    },
                    "attention": {
                        "normal": "#cc3300",
                        "subtle": "#DDcc3300"
                    }
                },
                "backgroundColor": "#FFFFFF"
            },
            "emphasis": {
                "fontColors": {
                    "default": {
                        "normal": "#333333",
                        "subtle": "#EE333333"
                    },
                    "accent": {
                        "normal": "#2E89FC",
                        "subtle": "#882E89FC"
                    },
                    "good": {
                        "normal": "#54a254",
                        "subtle": "#DD54a254"
                    },
                    "warning": {
                        "normal": "#e69500",
                        "subtle": "#DDe69500"
                    },
                    "attention": {
                        "normal": "#cc3300",
                        "subtle": "#DDcc3300"
                    }
                },
                "backgroundColor": "#08000000"
            }
        },
        "spacing": {
            "small": 3,
            "default": 8,
            "medium": 20,
            "large": 30,
            "extraLarge": 40,
            "padding": 20
        },
        "separator": {
            "lineThickness": 1,
            "lineColor": "#EEEEEE"
        },
        "actions": {
            "maxActions": 5,
            "spacing": "Default",
            "buttonSpacing": 10,
            "showCard": {
                "actionMode": "Popup",
                "inlineTopMargin": 16,
                "style": "Emphasis"
            },
            "preExpandSingleShowCardAction": false,
            "actionsOrientation": "Horizontal",
            "actionAlignment": "Left"
        },
        "adaptiveCard": {
            "allowCustomStyle": false
        },
        "imageSet": {
            "imageSize": "Medium",
            "maxImageHeight": "maxImageHeight"
        },
        "factSet": {
            "title": {
                "size": "Default",
                "color": "Default",
                "isSubtle": false,
                "weight": "Bolder",
                "warp": true
            },
            "value": {
                "size": "Default",
                "color": "Default",
                "isSubtle": false,
                "weight": "Default",
                "warp": true
            },
            "spacing": 10
        }
    };

    this.summary = json["summary"];
    this.themeColor = json["themeColor"];
    if (json["style"]) {
        this.style = json["style"];
    }

    this._adaptiveCard = new AdaptiveCards.AdaptiveCard();
    this._adaptiveCard.hostConfig = new AdaptiveCards.HostConfig(this.defaultCardConfig);

    if (json["title"] != undefined) {
        var textBlock = new AdaptiveCards.TextBlock();
        textBlock.text = json["title"];
        textBlock.size = "large";
        textBlock.wrap = true;
        this._adaptiveCard.addItem(textBlock);
    }

    if (json["text"] != undefined) {
        var textBlock = new AdaptiveCards.TextBlock();
        textBlock.text = json["text"],
        textBlock.wrap = true;
        this._adaptiveCard.addItem(textBlock);
    }

    if (json["sections"] != undefined) {
        var sectionArray = json["sections"];
        for (var i = 0; i < sectionArray.length; i++) {
            var section = parseSection(sectionArray[i], this.hostContainer);
            this._adaptiveCard.addItem(section);
        }
    }
    if (json["potentialAction"] != undefined) {
        var actionSet = parseActionSet(json["potentialAction"], this.hostContainer);
        actionSet.actionStyle = "link";
        this._adaptiveCard.addItem(actionSet);
    }
};
MessageCard.prototype.render = function () {
    return this._adaptiveCard.render();
};

function parsePicture(json, defaultSize, defaultStyle) {
    if (defaultSize === void 0) { defaultSize = "medium"; }
    if (defaultStyle === void 0) { defaultStyle = "normal"; }
    var picture = new AdaptiveCards.Image();
    picture.url = json["image"];
    picture.size = json["size"] ? json["size"] : defaultSize;
    return picture;
}

function parseImageSet(json) {
    var imageSet = new AdaptiveCards.ImageSet();
    var imageArray = json;
    for (var i = 0; i < imageArray.length; i++) {
        var image = parsePicture(imageArray[i], "small");
        imageSet.addImage(image);
    }
    return imageSet;
}

function parseFactSet(json) {
    var factSet = new AdaptiveCards.FactSet();
    var factArray = json;
    for (var i = 0; i < factArray.length; i++) {
        var fact = new AdaptiveCards.Fact();
        fact.name = factArray[i]["name"];
        fact.value = factArray[i]["value"];
        factSet.facts.push(fact);
    }
    return factSet;
}

function parseOpenUrlAction(json) {
    var action = new AdaptiveCards.OpenUrlAction();
    action.title = json["name"];
    action.url = "TODO";
    return action;
}

function parseHttpAction(json) {
    var mobileRender = new AdaptiveCardMobileRender();
    var action = new mobileRender.HttpAction();
    action.method = "POST";
    action.body = json["body"];
    action.title = json["name"];
    action.url = json["url"];
    return action;
}

function parseInvokeAddInCommandAction(json) {
    var action = new InvokeAddInCommandAction();
    action.title = json["name"];
    action.addInId = json["addInId"];
    action.desktopCommandId = json["desktopCommandId"];
    action.initializationContext = json["initializationContext"];
    return action;
}

function parseInput(input, json) {
    input.id = json["id"];
    input.defaultValue = json["value"];
}

function parseTextInput(json) {
    var input = new AdaptiveCards.TextInput();
    parseInput(input, json);
    input.placeholder = json["title"];
    input.isMultiline = json["isMultiline"];
    return input;
}

function parseDateInput(json) {
    var input = new AdaptiveCards.DateInput();
    parseInput(input, json);
    return input;
}

function parseChoiceSetInput(json) {
    var input = new AdaptiveCards.ChoiceSetInput();
    parseInput(input, json);
    input.placeholder = json["title"];
    var choiceArray = json["choices"];
    if (choiceArray) {
        for (var i = 0; i < choiceArray.length; i++) {
            var choice = new AdaptiveCards.Choice();
            choice.title = choiceArray[i]["display"];
            choice.value = choiceArray[i]["value"];
            input.choices.push(choice);
        }
    }
    input.isMultiSelect = json["isMultiSelect"];
    input.isCompact = !(json["style"] === "expanded");
    return input;
}

function parseShowCardAction(json, host) {
    var showCardAction = new AdaptiveCards.ShowCardAction();
    showCardAction.title = json["name"];
    showCardAction.card.actionStyle = "button";
    var inputArray = json["inputs"];
    if (inputArray) {
        for (var i = 0; i < inputArray.length; i++) {
            var jsonInput = inputArray[i];
            var input = null;
            switch (jsonInput["@type"]) {
                case "TextInput":
                    input = parseTextInput(jsonInput);
                    break;
                case "DateInput":
                    input = parseDateInput(jsonInput);
                    break;
                case "MultichoiceInput":
                    input = parseChoiceSetInput(jsonInput);
                    break;
            }
            if (input) {
                showCardAction.card.addItem(input);
            }
        }
    }
    var actionArray = json["actions"];
    if (actionArray) {
        showCardAction.card.addItem(parseActionSet(actionArray, host));
    }
    return showCardAction;
}

function parseActionSet(json, host) {
    var actionSet = new AdaptiveCards.ActionSet();
    var actionArray = json;
    for (var i = 0; i < actionArray.length; i++) {
        var jsonAction = actionArray[i];
        var action = null;
        switch (jsonAction["@type"]) {
            case "OpenUri":
                action = parseOpenUrlAction(jsonAction);
                break;
            case "HttpPOST":
                action = parseHttpAction(jsonAction);
                break;
            case "InvokeAddInCommand":
                action = parseInvokeAddInCommandAction(jsonAction);
                break;
            case "ActionCard":
                if (host.allowActionCard) {
                    action = parseShowCardAction(jsonAction, host);
                }
                break;
        }
        if (action) {
            actionSet.addAction(action);
        }
    }
    return actionSet;
}

function parseSection(json, host) {
    var section = new AdaptiveCards.Container();
    section.separation = json["startGroup"] ? "strong" : "default";
    if (json["title"] != undefined) {
        var textBlock = new AdaptiveCards.TextBlock();
        textBlock.text = json["title"];
        textBlock.size = "medium";
        textBlock.wrap = true;
        section.addItem(textBlock);
    }
    if(json["style"] != null)
    {
        section.style = json["style"] == "emphasis" ? "emphasis" : "normal";
    }
    if (json["activityTitle"] != undefined || json["activitySubtitle"] != undefined ||
        json["activityText"] != undefined || json["activityImage"] != undefined) {
        var columnSet = new AdaptiveCards.ColumnSet();
        var column;
        // Image column
        if (json["activityImage"] != null) {
            column = new AdaptiveCards.Column();
            column.size = "auto";
            var image = new AdaptiveCards.Image();
            image.size = json["activityImageSize"] ? json["activityImageSize"] : "small";
            image.style = json["activityImageStyle"] ? json["activityImageStyle"] : "person";
            image.url = json["activityImage"];
            column.addItem(image);
            columnSet.addColumn(column);
        }
        // Text column
        column = new AdaptiveCards.Column;
        column.size = "stretch";
        if (json["activityTitle"] != null) {
            var textBlock_1 = new AdaptiveCards.TextBlock();
            textBlock_1.text = json["activityTitle"];
            textBlock_1.separation = "none";
            textBlock_1.wrap = true;
            column.addItem(textBlock_1);
        }
        if (json["activitySubtitle"] != null) {
            var textBlock_2 = new AdaptiveCards.TextBlock();
            textBlock_2.text = json["activitySubtitle"];
            textBlock_2.weight = "lighter";
            textBlock_2.isSubtle = true;
            textBlock_2.separation = "none";
            textBlock_2.wrap = true;
            column.addItem(textBlock_2);
        }
        if (json["activityText"] != null) {
            var textBlock_3 = new AdaptiveCards.TextBlock();
            textBlock_3.text = json["activityText"];
            textBlock_3.separation = "none";
            textBlock_3.wrap = true;
            column.addItem(textBlock_3);
        }
        columnSet.addColumn(column);
        section.addItem(columnSet);
    }
    if (host.allowHeroImage) {
        var heroImage = json["heroImage"];
        if (heroImage != undefined) {
            var image_1 = parsePicture(heroImage);
            image_1.size = "auto";
            section.addItem(image_1);
        }
    }
    if (json["text"] != undefined) {
        var text = new AdaptiveCards.TextBlock();
        text.text = json["text"];
        text.wrap = true;
        section.addItem(text);
    }
    if (host.allowFacts) {
        if (json["facts"] != undefined) {
            var factGroup = parseFactSet(json["facts"]);
            section.addItem(factGroup);
        }
    }
    if (host.allowImages) {
        if (json["images"] != undefined) {
            var pictureGallery = parseImageSet(json["images"]);
            section.addItem(pictureGallery);
        }
    }
    if (json["potentialAction"] != undefined) {
        var actionSet = parseActionSet(json["potentialAction"], host);
        actionSet.actionStyle = "link";
        section.addItem(actionSet);
    }
    console.log(section);
    return section;
}

window.renderCard = function(json) {
  var cardRenderer = new AdaptiveCardMobileRender();
  var parsedJSON = JSON.parse(window.atob(json));
  cardRenderer.render(parsedJSON);
}
