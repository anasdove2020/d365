// CONSTANTS VARIABLE

const FIELD_NAME_QUOTE_NUMBER = "new_quotenumber";
const FIELD_NAME_PURCHASE_ORDER_NUMBER = "new_purchaseordernumber";
const FIELD_NAME_ORDER_NUMBER = "new_ordernumber";
const FIELD_NAME_STATUS_CODE = "statuscode";
const FIELD_NAME_ASSIGN_TO = "ownerid";
const FIELD_NAME_DUE_DATE = "new_duedate";

const ENQUIRY_PHASE = "Enquiry"; 
const QUOTE_PHASE = "Quote"; 
const ORDER_PHASE = "Order";
const DRAW_PHASE = "Draw";
const END_OF_LIFE_PHASE = "End Of Life";

const NEW_STATUS_CODE = 1;
const QUOTE_SENT_STATUS_CODE = 2;
const ORDER_RECEIVED_STATUS_CODE = 3;
const DRAWING_IN_PROGRESS_STATUS_CODE = 4;
const DRAWN_STATUS_CODE = 675430001;
const NOT_WON_STATUS_CODE = 100000001;

// TODO: as current release only for Plate and Steel, in future need to modify to cover teams from different business units.
const SALES_TEAM_NAME = "Plate And Steel Sales Team"; 

// Initial Form Variable

let currentQuoteNumber = null;
let currentPurchaseOrderNumber = null;
let currentOrderNumber = null;
let currentStatusCode = 0;

// -------------------------------------------------------------------------------

window.caseDetailOnLoad = function(executionContext) {
    const formContext = executionContext.getFormContext();

    const interval = setInterval(() => {
        try {
            const globalDoc = window.top.document;

            const btn = globalDoc.querySelector('button[aria-label="Teams chats"]');

            if (btn) {
                if (btn.getAttribute("aria-expanded") === "true") {
                    btn.click();
                }

                clearInterval(interval);
            }
        } catch (e) {
            console.error("❌ Error accessing DOM:", e);
        }
    }, 1000);

    setTimeout(() => clearInterval(interval), 30000);

    initCaseFormField(formContext);
    
    applyDueDate(formContext);
}

// SAP-51 Move Due Date Validation from Enquiry Phase to Order Phase
function applyDueDate(formContext){
	if(isCurrentPhaseOrder(formContext) || isCurrentPhaseDraw(formContext) || isCurrentPhaseEndOfLife(formContext))
		setDueDateRequired(formContext);
}

// -------------------------------------------------------------------------------

window.quoteNumberOnChange = function(executionContext) {
    var formContext = executionContext.getFormContext();
    
    var quoteNumber = getFieldNameValue(formContext, FIELD_NAME_QUOTE_NUMBER);

    clearFieldNotification(formContext, FIELD_NAME_QUOTE_NUMBER);

    if (isAllowToSetQuoteNumber(formContext, quoteNumber)) {
        setStatusCodeQuoteSent(formContext);
        saveDataEntity(formContext);
        currentQuoteNumber = quoteNumber;
		
		applyDueDate(formContext);

        return;
    }

    if (isAllowToClearQuoteNumber(formContext, quoteNumber)) {
        setStatusCodeNew(formContext);
        saveDataEntity(formContext);
        currentQuoteNumber = quoteNumber;
		
		applyDueDate(formContext);
		
        return;
    }

    if (isBlank(quoteNumber)) {
        revertQuoteNumber(formContext);

        const message = `You are not allowed to remove Quote Number in ${getCurrentPhase(formContext)} phase.`;
        setFieldNotification(formContext, FIELD_NAME_QUOTE_NUMBER, message);

        setTimeout(() => {
            clearFieldNotification(formContext, FIELD_NAME_QUOTE_NUMBER);
        }, 5000);
    } else {
        saveDataEntity(formContext);
        currentQuoteNumber = quoteNumber;
		
		applyDueDate(formContext);
    }
}

// -------------------------------------------------------------------------------

window.assignToOnChange = function(executionContext) {
    const formContext = executionContext.getFormContext();
    const user = isAssignToEntityTypeSystemUser(formContext);

    if (user) {
        const userId = getUserId(user);

        if (isCurrentPhaseEnquiry(formContext)) {
            clearErrorFormNotification(formContext);

            checkTeamMembership(
                formContext, 
                SALES_TEAM_NAME, 
                userId, 
                () => WhenUserBelongToSalesTeam(
                    formContext, () => nextPhaseFromEnquiryToQuote(formContext)
                )
            );
        } else {
            saveDataEntity(formContext);
        }
    } else {
        const salesTeam = isAssignToEntityTypeTeamAndSalesTeam(formContext);

        if (salesTeam) {
            if (isCurrentPhaseQuote(formContext)) {
                saveData(formContext, () => previousPhaseFromQuoteToEnquiry(formContext));
            } else {
                saveDataEntity(formContext);
            }
        } else {
            saveDataEntity(formContext);
        }
    }
}

// -------------------------------------------------------------------------------

window.purchaseOrderNumberOnChange = function(executionContext) {
    const formContext = executionContext.getFormContext();

    var quoteNumber = getFieldNameValue(formContext, FIELD_NAME_QUOTE_NUMBER);
    var purchaseOrderNumber = getFieldNameValue(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);

    clearFieldNotification(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);

    if (isBlank(quoteNumber)) {
        revertPurchaseOrderNumber(formContext);

        const message = "You must enter a Quote Number before adding a Purchase Order Number."
        setFieldNotification(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER, message);

        setTimeout(() => {
            clearFieldNotification(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);
        }, 5000);

        return;
    }

    if (isAllowToSetPurchaseOrderNumber(formContext, purchaseOrderNumber)) {
        setStatusCodeOrderReceived(formContext);
        saveData(formContext, () => nextPhaseFromQuoteToOrder(formContext));
        currentPurchaseOrderNumber = purchaseOrderNumber;
        currentStatusCode = ORDER_RECEIVED_STATUS_CODE;
        return;
    }

    if (isAllowToClearPurchaseOrderNumber(formContext, purchaseOrderNumber)) {
        setStatusCodeQuoteSent(formContext);
        saveData(formContext, () => previousPhaseFromOrderToQuote(formContext));
        currentPurchaseOrderNumber = purchaseOrderNumber;
        currentStatusCode = QUOTE_SENT_STATUS_CODE;
        return;
    }
    
    if (isBlank(purchaseOrderNumber)) {
        if (isCurrentPhaseEnquiry(formContext) || isCurrentPhaseQuote(formContext)) {
            saveDataEntity(formContext);
            currentPurchaseOrderNumber = purchaseOrderNumber;
        } else {
            revertPurchaseOrderNumber(formContext);

            const message = `You are not allowed to remove Purchase Order Number in ${getCurrentPhase(formContext)} phase.`;
            setFieldNotification(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER, message);

            setTimeout(() => {
                clearFieldNotification(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);
            }, 5000);
        }
    } else {
        saveDataEntity(formContext);
        currentPurchaseOrderNumber = purchaseOrderNumber;
    }
}

// -------------------------------------------------------------------------------

window.orderNumberOnChange = function(executionContext) {
    const formContext = executionContext.getFormContext();

    var quoteNumber = getFieldNameValue(formContext, FIELD_NAME_QUOTE_NUMBER);
    var purchaseOrderNumber = getFieldNameValue(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);
    var orderNumber = getFieldNameValue(formContext, FIELD_NAME_ORDER_NUMBER);
    
    var dueDate = getFieldNameValue(formContext, FIELD_NAME_DUE_DATE);

    clearFieldNotification(formContext, FIELD_NAME_ORDER_NUMBER);

    if (isBlank(quoteNumber) && isBlank(purchaseOrderNumber)) {
        revertOrderNumber(formContext);

        const message = "You must enter both Quote Number and Purchase Order Number before adding an Order Number."
        setFieldNotification(formContext, FIELD_NAME_ORDER_NUMBER, message);

        setTimeout(() => {
            clearFieldNotification(formContext, FIELD_NAME_ORDER_NUMBER);
        }, 5000);

        return;
    }

    if (isBlank(purchaseOrderNumber)) {
        revertOrderNumber(formContext);

        const message = "You must enter Purchase Order Number before adding an Order Number."
        setFieldNotification(formContext, FIELD_NAME_ORDER_NUMBER, message);

        setTimeout(() => {
            clearFieldNotification(formContext, FIELD_NAME_ORDER_NUMBER);
        }, 5000);

        return;
    }

    // SAP-51 Move to Order Phase only occur if order number and due data is populated
    if (isAllowToSetOrderNumber(formContext, orderNumber, dueDate)) {
        setStatusCodeDrawingInProgress(formContext);
        saveData(formContext, () => nextPhaseFromOrderToDraw(formContext));
        currentOrderNumber = orderNumber;
        currentStatusCode = DRAWING_IN_PROGRESS_STATUS_CODE;
        return;
    }

    if (isAllowToClearOrderNumber(formContext, orderNumber)) {
        setStatusCodeOrderReceived(formContext);
        saveData(formContext, () => previousPhaseFromDrawToOrder(formContext));
        currentOrderNumber = orderNumber;
        currentStatusCode = ORDER_RECEIVED_STATUS_CODE;
        return;
    }
    
    if (isBlank(orderNumber)) {
        if (isCurrentPhaseEnquiry(formContext) || isCurrentPhaseQuote(formContext) || isCurrentPhaseOrder(formContext)) {
            saveDataEntity(formContext);
            currentOrderNumber = orderNumber;
        } else {
            revertOrderNumber(formContext);

            const message = `You are not allowed to remove Order Number in ${getCurrentPhase(formContext)} phase.`;
            setFieldNotification(formContext, FIELD_NAME_ORDER_NUMBER, message);

            setTimeout(() => {
                clearFieldNotification(formContext, FIELD_NAME_ORDER_NUMBER);
            }, 5000);
        }
    } else {
        saveDataEntity(formContext);
        currentOrderNumber = orderNumber;
    }
}

// -------------------------------------------------------------------------------

window.statusCodeOnChange = function(executionContext) {
    const formContext = executionContext.getFormContext();

    var quoteNumber = getFieldNameValue(formContext, FIELD_NAME_QUOTE_NUMBER);
    var purchaseOrderNumber = getFieldNameValue(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);
    var orderNumber = getFieldNameValue(formContext, FIELD_NAME_ORDER_NUMBER);
    var statusCode = getFieldNameValue(formContext, FIELD_NAME_STATUS_CODE);

    if (isAllowToSetStatusNotWon(formContext)) {
        setDefaultQuoteNumberWhenIsBlank(formContext);
        setDefaultPurchaseOrderNumberWhenIsBlank(formContext);
        setDefaultOrderNumberWhenIsBlank(formContext);

        saveData(formContext, () => moveDirectlyToEndOfLife(formContext));
        currentStatusCode = statusCode;
        return;
    }

    if (isAllowToSetStatusDrawn(formContext, quoteNumber, purchaseOrderNumber, orderNumber)) {
        saveData(formContext, () => nextPhaseFromDrawToEndOfLife(formContext));
        currentStatusCode = statusCode;
        return;
    }

    if (isAllowToSetStatusDrawingInProgress(formContext, quoteNumber, purchaseOrderNumber, orderNumber)) {
        saveData(formContext, () => previousPhaseFromEndOfLifeToDraw(formContext));
        currentStatusCode = statusCode;
        return;
    }
    
    saveDataEntity(formContext);
    currentStatusCode = statusCode;
}

// -------------------------------------------------------------------------------

window.dueDateOnChange = function(executionContext) {
    const formContext = executionContext.getFormContext();
    var dueDate = getFieldNameValue(formContext, FIELD_NAME_DUE_DATE);
    var orderNumber = getFieldNameValue(formContext, FIELD_NAME_ORDER_NUMBER);

    // SAP-51 Move to Order Phase only occur if order number and due data is populated
    if (isAllowToSetOrderNumber(formContext, orderNumber, dueDate)) {
        setStatusCodeDrawingInProgress(formContext);
        saveData(formContext, () => nextPhaseFromOrderToDraw(formContext));
        currentOrderNumber = orderNumber;
        currentStatusCode = DRAWING_IN_PROGRESS_STATUS_CODE;
        return;
    }
}

// -------------------------------------------------------------------------------

// INIT CASE FORM FIELD

function initCaseFormField(formContext) {
    currentQuoteNumber = getFieldNameValue(formContext, FIELD_NAME_QUOTE_NUMBER);
    currentPurchaseOrderNumber = getFieldNameValue(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);
    currentOrderNumber = getFieldNameValue(formContext, FIELD_NAME_ORDER_NUMBER);
    currentStatusCode = getFieldNameValue(formContext, FIELD_NAME_STATUS_CODE);
}

function setDueDateRequired(formContext) {
    const dueDateAttr = getFieldNameAttr(formContext, FIELD_NAME_DUE_DATE);
    if (dueDateAttr) {
        dueDateAttr.setRequiredLevel("required");
    }
}

// IS ALLOW TO SET QUOTE NUMBER

function isAllowToSetQuoteNumber(formContext, quoteNumber) {
    return isCurrentStatusNew(formContext) &&
           isNotBlank(quoteNumber);
}

// IS ALLOW TO CLEAR QUOTE NUMBER

function isAllowToClearQuoteNumber(formContext, quoteNumber) {
    return isCurrentPhaseEnquiry(formContext) &&
           isCurrentStatusQuoteSent(formContext) &&
           isBlank(quoteNumber);
}

// IS ALLOW TO SET PURCHASE ORDER NUMBER

function isAllowToSetPurchaseOrderNumber(formContext, purchaseOrderNumber) {
    return isCurrentPhaseQuote(formContext) &&
           isCurrentStatusQuoteSent(formContext) &&
           isNotBlank(purchaseOrderNumber);
}

// IS ALLOW TO CLEAR PURCHASE ORDER NUMBER

function isAllowToClearPurchaseOrderNumber(formContext, purchaseOrderNumber) {
    return isCurrentPhaseOrder(formContext) &&
           isCurrentStatusOrderReceived(formContext) &&
           isBlank(purchaseOrderNumber);
}

// IS ALLOW TO SET ORDER NUMBER

function isAllowToSetOrderNumber(formContext, orderNumber, dueDate) {
    return isCurrentPhaseOrder(formContext) &&
           isCurrentStatusOrderReceived(formContext) &&
           isNotBlank(orderNumber) && isNotBlank(dueDate);
}

// IS ALLOW TO CLEAR ORDER NUMBER

function isAllowToClearOrderNumber(formContext, orderNumber) {
    return isCurrentPhaseDraw(formContext) &&
           isCurrentStatusDrawingInProgress(formContext) &&
           isBlank(orderNumber);
}

// IS ALLOW TO SET STATUS DRAWN

function isAllowToSetStatusDrawn(formContext, quoteNumber, purchaseOrderNumber, orderNumber) {
    return isCurrentPhaseDraw(formContext) &&
           isCurrentStatusDrawn(formContext) &&
           isCurrentInitialStatusDrawingInProgress(formContext) &&
           isNotBlank(quoteNumber) &&
           isNotBlank(purchaseOrderNumber) &&
           isNotBlank(orderNumber);
}

// IS ALLOW TO SET STATUS DRAWING IN PROGRESS

function isAllowToSetStatusDrawingInProgress(formContext, quoteNumber, purchaseOrderNumber, orderNumber) {
    return isCurrentPhaseEndOfLife(formContext) &&
           isCurrentStatusDrawingInProgress(formContext) &&
           isCurrentInitialStatusDrawn(formContext) &&
           isNotBlank(quoteNumber) &&
           isNotBlank(purchaseOrderNumber) &&
           isNotBlank(orderNumber);
}

// IS ALLOW TO SET STATUS NOT WON

function isAllowToSetStatusNotWon(formContext) {
    return isCurrentStatusNotWon(formContext);
}

// IS NOT BLANK

function isNotBlank(value) {
    return value && value.toString().trim() !== "";
}

// IS BLANK

function isBlank(value) {
    return !isNotBlank(value);
}

// GET FIELD NAME ATTRIBUTE

function getFieldNameAttr(formContext, fieldName) {
    return formContext.getAttribute(fieldName);
}

// GET FIELD NAME VALUE

function getFieldNameValue(formContext, fieldName) {
    var fieldNameAttr = getFieldNameAttr(formContext, fieldName);
    if (!fieldNameAttr) return null;

    const fieldNameValue = fieldNameAttr.getValue();
    return fieldNameValue;
}

// GET FIELD NAME INITIAL VALUE

function getFieldNameInitialValue(formContext, fieldName) {
    var fieldNameAttr = getFieldNameAttr(formContext, fieldName);
    if (!fieldNameAttr) return null;

    const fieldNameInitialValue = fieldNameAttr.getInitialValue();
    return fieldNameInitialValue;
}

// REVERT QUOTE NUMBER

function revertQuoteNumber(formContext) {
    var quoteNumberAttr = getFieldNameAttr(formContext, FIELD_NAME_QUOTE_NUMBER);
    if (!quoteNumberAttr) return;

    quoteNumberAttr.setValue(currentQuoteNumber);
}

// REVERT PURCHASE ORDER NUMBER

function revertPurchaseOrderNumber(formContext) {
    var purchaseOrderNumberAttr = getFieldNameAttr(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);
    if (!purchaseOrderNumberAttr) return;

    purchaseOrderNumberAttr.setValue(currentPurchaseOrderNumber);
}

// REVERT ORDER NUMBER

function revertOrderNumber(formContext) {
    var orderNumberAttr = getFieldNameAttr(formContext, FIELD_NAME_ORDER_NUMBER);
    if (!orderNumberAttr) return;

    orderNumberAttr.setValue(currentOrderNumber);
}

// REVERT STATUS CODE

function revertStatusCode(formContext) {
    var statusCodeAttr = getFieldNameAttr(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeAttr) return;

    statusCodeAttr.setValue(currentStatusCode);
    statusCodeAttr.setSubmitMode("always");
}

// SET DEFAULT QUOTE NUMBER WITH DEFAULT VALUE IF BLANK

function setDefaultQuoteNumberWhenIsBlank(formContext) {
    var quoteNumberAttr = getFieldNameAttr(formContext, FIELD_NAME_QUOTE_NUMBER);
    if (!quoteNumberAttr) return;
    
    var quoteNumberValue = getFieldNameValue(formContext, FIELD_NAME_QUOTE_NUMBER);
    if (!quoteNumberValue) {
        quoteNumberAttr.setValue('-');
        quoteNumberAttr.setSubmitMode("always");
    }
}

// SET DEFAULT PURCHASE ORDER NUMBER WITH DEFAULT VALUE IF BLANK

function setDefaultPurchaseOrderNumberWhenIsBlank(formContext) {
    var purchaseOrderNumberAttr = getFieldNameAttr(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);
    if (!purchaseOrderNumberAttr) return;
    
    var purchaseOrderNumberValue = getFieldNameValue(formContext, FIELD_NAME_PURCHASE_ORDER_NUMBER);
    if (!purchaseOrderNumberValue) {
        purchaseOrderNumberAttr.setValue('-');
        purchaseOrderNumberAttr.setSubmitMode("always");
    }
}

// SET DEFAULT ORDER NUMBER WITH DEFAULT VALUE IF BLANK

function setDefaultOrderNumberWhenIsBlank(formContext) {
    var orderNumberAttr = getFieldNameAttr(formContext, FIELD_NAME_ORDER_NUMBER);
    if (!orderNumberAttr) return;
    
    var orderNumberValue = getFieldNameValue(formContext, FIELD_NAME_ORDER_NUMBER);
    if (!orderNumberValue) {
        orderNumberAttr.setValue('-');
        orderNumberAttr.setSubmitMode("always");
    }
}

// SET FIELD NOTIFICATION

function setFieldNotification(formContext, fieldName, message) {
    var fieldNameControl = formContext.getControl(fieldName);
    if (!fieldNameControl) return;

    fieldNameControl.setNotification(message);
}

// CLEAR FIELD NOTIFICATION

function clearFieldNotification(formContext, fieldName) {
    var fieldNameControl = formContext.getControl(fieldName);
    if (!fieldNameControl) return;

    fieldNameControl.clearNotification();
}

// SAVE DATA ENTITY

function saveDataEntity(formContext) {
    formContext.data.entity.save();
}

// SAVE DATA

function saveData(formContext, saveSuccessFunc) {
    formContext.data.save().then(
        () => {
            saveSuccessFunc(formContext);
        },
        error => {
            const message = "Failed to save data: " + error.message;
            setErrorFormNotification(formContext, message);
        }
    );
}

// IS CURRENT STATUS CODE

function isCurrentStatusNew(formContext) {
    var statusCodeValue = getFieldNameValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === NEW_STATUS_CODE;
}

function isCurrentStatusQuoteSent(formContext) {
    var statusCodeValue = getFieldNameValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === QUOTE_SENT_STATUS_CODE;
}

function isCurrentStatusOrderReceived(formContext) {
    var statusCodeValue = getFieldNameValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === ORDER_RECEIVED_STATUS_CODE;
}

function isCurrentStatusDrawingInProgress(formContext) {
    var statusCodeValue = getFieldNameValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === DRAWING_IN_PROGRESS_STATUS_CODE;
}

function isCurrentStatusDrawn(formContext) {
    var statusCodeValue = getFieldNameValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === DRAWN_STATUS_CODE;
}

function isCurrentStatusNotWon(formContext) {
    var statusCodeValue = getFieldNameValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === NOT_WON_STATUS_CODE;
}

// IS CURRENT INITIAL STATUS CODE

function isCurrentInitialStatusNew(formContext) {
    var statusCodeValue = getFieldNameInitialValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === NEW_STATUS_CODE;
}

function isCurrentInitialStatusQuoteSent(formContext) {
    var statusCodeValue = getFieldNameInitialValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === QUOTE_SENT_STATUS_CODE;
}

function isCurrentInitialStatusOrderReceived(formContext) {
    var statusCodeValue = getFieldNameInitialValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === ORDER_RECEIVED_STATUS_CODE;
}

function isCurrentInitialStatusDrawingInProgress(formContext) {
    var statusCodeValue = getFieldNameInitialValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === DRAWING_IN_PROGRESS_STATUS_CODE;
}

function isCurrentInitialStatusDrawn(formContext) {
    var statusCodeValue = getFieldNameInitialValue(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeValue) return;

    return statusCodeValue === DRAWN_STATUS_CODE;
}

// SET STATUS CODE

function setStatusCodeNew(formContext) {
    var statusCodeAttr = getFieldNameAttr(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeAttr) return;

    statusCodeAttr.setValue(NEW_STATUS_CODE);
    statusCodeAttr.setSubmitMode("always");
}

function setStatusCodeQuoteSent(formContext) {
    var statusCodeAttr = getFieldNameAttr(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeAttr) return;

    statusCodeAttr.setValue(QUOTE_SENT_STATUS_CODE);
    statusCodeAttr.setSubmitMode("always");
}

function setStatusCodeOrderReceived(formContext) {
    var statusCodeAttr = getFieldNameAttr(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeAttr) return;

    statusCodeAttr.setValue(ORDER_RECEIVED_STATUS_CODE);
    statusCodeAttr.setSubmitMode("always");
}

function setStatusCodeDrawingInProgress(formContext) {
    var statusCodeAttr = getFieldNameAttr(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeAttr) return;

    statusCodeAttr.setValue(DRAWING_IN_PROGRESS_STATUS_CODE);
    statusCodeAttr.setSubmitMode("always");
}

function setStatusCodeDrawn(formContext) {
    var statusCodeAttr = getFieldNameAttr(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeAttr) return;

    statusCodeAttr.setValue(DRAWN_STATUS_CODE);
    statusCodeAttr.setSubmitMode("always");
}

function setStatusCodeNotWon(formContext) {
    var statusCodeAttr = getFieldNameAttr(formContext, FIELD_NAME_STATUS_CODE);
    if (!statusCodeAttr) return;

    statusCodeAttr.setValue(DRAWN_STATUS_CODE);
    statusCodeAttr.setSubmitMode("always");
}

// PHASE

function getProcess(formContext) {
    return formContext.data.process;
}

function getCurrentPhase(formContext) {
    var process = getProcess(formContext);
    var activeStage = process.getActiveStage();
    if (!activeStage) return null;
    return activeStage.getName();
}

// IS CURRENT PHASE

function isCurrentPhaseEnquiry(formContext) {
    var currentPhase = getCurrentPhase(formContext);
    if (!currentPhase) return false;
    return currentPhase === ENQUIRY_PHASE;
}

function isCurrentPhaseQuote(formContext) {
    var currentPhase = getCurrentPhase(formContext);
    if (!currentPhase) return false;
    return currentPhase === QUOTE_PHASE;
}

function isCurrentPhaseOrder(formContext) {
    var currentPhase = getCurrentPhase(formContext);
    if (!currentPhase) return false;
    return currentPhase === ORDER_PHASE;
}

function isCurrentPhaseDraw(formContext) {
    var currentPhase = getCurrentPhase(formContext);
    if (!currentPhase) return false;
    return currentPhase === DRAW_PHASE;
}

function isCurrentPhaseEndOfLife(formContext) {
    var currentPhase = getCurrentPhase(formContext);
    if (!currentPhase) return false;
    return currentPhase === END_OF_LIFE_PHASE;
}

// IS ASSIGN TO ENTITY TYPE SYSTEM USER

function isAssignToEntityTypeSystemUser(formContext) {
    var assignToValue = getFieldNameValue(formContext, FIELD_NAME_ASSIGN_TO);

    if (!assignToValue) return null;

    if (assignToValue.length === 0) return null;

    if (assignToValue[0].entityType.toLowerCase() !== "systemuser") return null;

    return assignToValue[0];
}

// IS ASSIGN TO ENTITY TYPE TEAM

function isAssignToEntityTypeTeam(formContext) {
    var assignToValue = getFieldNameValue(formContext, FIELD_NAME_ASSIGN_TO);

    if (!assignToValue) return null;

    if (assignToValue.length === 0) return null;

    if (assignToValue[0].entityType.toLowerCase() !== "team") return null;

    return assignToValue[0];
}

// IS ASSIGN TO ENTITY TYPE TEAM AND SALES TEAM

function isAssignToEntityTypeTeamAndSalesTeam(formContext) {
    const isTeam = isAssignToEntityTypeTeam(formContext);

    if (!isTeam) return null;

    if (isTeam.name === SALES_TEAM_NAME) {
        return true;
    }

    return false;
}

// RESET ASSIGN TO VALUE

function resetAssignToValue(formContext) {
    var assignToAttr = getFieldNameAttr(formContext, FIELD_NAME_ASSIGN_TO);
    assignToAttr.setValue(null);
}

// SET INFO FORM NOTIFICATION

function setInfoFormNotification(formContext, message) {
    formContext.ui.setFormNotification(message, "INFO", "infoFormNotification");
}

// SET ERROR FORM NOTIFICATION

function setErrorFormNotification(formContext, message) {
    formContext.ui.setFormNotification(message, "ERROR", "errorFormNotification");
}

// CLEAR INFO FORM NOTIFICATION

function clearInfoFormNotification(formContext) {
    formContext.ui.clearFormNotification("infoFormNotification");
}

// CLEAR ERROR FORM NOTIFICATION

function clearErrorFormNotification(formContext) {
    formContext.ui.clearFormNotification("errorFormNotification");
}

// GET USER ID

function getUserId(user) {
    return user.id.replace(/[{}]/g, "");
}

// CHECK TEAM MEMBERSHIP

function checkTeamMembership(formContext, teamName, userId, inTeamFunc) {
    Xrm.WebApi.retrieveMultipleRecords(
        "team",
        `?$select=teamid,name&$filter=name eq '${teamName}'` +
        `&$expand=teammembership_association($filter=systemuserid eq ${userId})`
    ).then(
        result => {
            if (
                result.entities.length > 0 &&
                result.entities[0]["teammembership_association"].length > 0
            ) {
                inTeamFunc(formContext);
            } else {
                resetAssignToValue(formContext);

                const errorMessage = `Selected user does not belong to ${teamName}.`;
                setErrorFormNotification(formContext, errorMessage);
            }
        },
        error => {
            console.error("Error checking team membership: " + error.message);
        }
    );
}

// WHEN USER BELONG TO SALES TEAM

function WhenUserBelongToSalesTeam(formContext, saveSuccessFunc) {
    if (isCurrentPhaseEnquiry(formContext)) {
        formContext.data.save().then(
            () => {
                saveSuccessFunc(formContext);
            },
            error => {
                const message = "Unable to save before moving phase: " + error.message;
                setErrorFormNotification(formContext, message);
            }
        );
    }
}

// NEXT PHASE

function nextPhase(formContext, current, next) {
    formContext.data.process.moveNext(res => {
        if (res === "success") {
            const message = `Phase successfully advanced from ${current} to ${next}.`;
            setInfoFormNotification(formContext, message);
			
			applyDueDate(formContext);
			
            setTimeout(() => {
                clearInfoFormNotification(formContext);
            }, 5000);
        } else {
            const message = `Failed to advance phase from ${current} to ${next}.`;
            setErrorFormNotification(formContext, message);
        }
    });
}

// NEXT PHASE FROM ENQUIRY TO QUOTE

function nextPhaseFromEnquiryToQuote(formContext) {
    nextPhase(formContext, ENQUIRY_PHASE, QUOTE_PHASE)
}

// NEXT PHASE FROM QUOTE TO ORDER

function nextPhaseFromQuoteToOrder(formContext) {
    nextPhase(formContext, QUOTE_PHASE, ORDER_PHASE)
}

// NEXT PHASE FROM ORDER TO DRAW

function nextPhaseFromOrderToDraw(formContext) {
    nextPhase(formContext, ORDER_PHASE, DRAW_PHASE)
}

// NEXT PHASE FROM DRAW TO END OF LIFE

function nextPhaseFromDrawToEndOfLife(formContext) {
    nextPhase(formContext, DRAW_PHASE, END_OF_LIFE_PHASE)
}

// PREVIOUS PHASE

function previousPhase(formContext, current, next) {
    formContext.data.process.movePrevious(res => {
        if (res === "success") {
            const message = `Phase successfully advanced from ${current} to ${next}.`;
            setInfoFormNotification(formContext, message);
            setTimeout(() => {
                clearInfoFormNotification(formContext);
            }, 5000);
        } else {
            const message = `Failed to advance phase from ${current} to ${next}.`;
            setErrorFormNotification(formContext, message);
        }
    });
}

// PREVIOUS PHASE FROM QUOTE TO ENQUIRY

function previousPhaseFromQuoteToEnquiry(formContext) {
    previousPhase(formContext, QUOTE_PHASE, ENQUIRY_PHASE)
}

// PREVIOUS PHASE FROM ORDER TO QUOTE

function previousPhaseFromOrderToQuote(formContext) {
    previousPhase(formContext, ORDER_PHASE, QUOTE_PHASE)
}

// PREVIOUS PHASE FROM DRAW TO ORDER

function previousPhaseFromDrawToOrder(formContext) {
    previousPhase(formContext, DRAW_PHASE, ORDER_PHASE)
}

// PREVIOUS PHASE FROM END OF LIFE TO DRAW 

function previousPhaseFromEndOfLifeToDraw(formContext) {
    previousPhase(formContext, END_OF_LIFE_PHASE, DRAW_PHASE)
}

// MOVE DIRECTLY TO END OF LIFE

function moveDirectlyToEndOfLife(formContext) {
    if (isCurrentPhaseEnquiry(formContext)) {
        // From Enquire to Quote
        formContext.data.process.moveNext(res => {
            if (res === "success") {
                // From Quote to Order
                formContext.data.process.moveNext(res => {
                    if (res === "success") {
                        // From Order to Draw
                        formContext.data.process.moveNext(res => {
                            if (res === "success") {
                                // From Draw to End Of Life
                                formContext.data.process.moveNext(res => {
                                    if (res === "success") {
                                        const message = `Phase successfully advanced from Enquiry to End Of Life.`;
                                        setInfoFormNotification(formContext, message);

                                        setTimeout(() => {
                                            clearInfoFormNotification(formContext);
                                        }, 5000);
                                    } else {
                                        const message = `Failed to advance phase from Draw to End Of Life.`;
                                        setErrorFormNotification(formContext, message);
                                    }
                                });
                            } else {
                                const message = `Failed to advance phase from Order to Draw.`;
                                setErrorFormNotification(formContext, message);
                            }
                        });
                    } else {
                        const message = `Failed to advance phase from Quote to Order.`;
                        setErrorFormNotification(formContext, message);
                    }
                });
            } else {
                const message = `Failed to advance phase from Enquiry to Quote.`;
                setErrorFormNotification(formContext, message);
            }
        });
    } else if (isCurrentPhaseQuote(formContext)) {
        // From Quote to Order
        formContext.data.process.moveNext(res => {
            if (res === "success") {
                // From Order to Draw
                formContext.data.process.moveNext(res => {
                    if (res === "success") {
                        // From Draw to End Of Life
                        formContext.data.process.moveNext(res => {
                            if (res === "success") {
                                const message = `Phase successfully advanced from Quote to End Of Life.`;
                                setInfoFormNotification(formContext, message);

                                setTimeout(() => {
                                    clearInfoFormNotification(formContext);
                                }, 5000);
                            } else {
                                const message = `Failed to advance phase from Draw to End Of Life.`;
                                setErrorFormNotification(formContext, message);
                            }
                        });
                    } else {
                        const message = `Failed to advance phase from Order to Draw.`;
                        setErrorFormNotification(formContext, message);
                    }
                });
            } else {
                const message = `Failed to advance phase from Quote to Order.`;
                setErrorFormNotification(formContext, message);
            }
        });
    } else if (isCurrentPhaseOrder(formContext)) {
        // From Order to Draw
        formContext.data.process.moveNext(res => {
            if (res === "success") {
                // From Draw to End Of Life
                formContext.data.process.moveNext(res => {
                    if (res === "success") {
                        const message = `Phase successfully advanced from Order to End Of Life.`;
                        setInfoFormNotification(formContext, message);

                        setTimeout(() => {
                            clearInfoFormNotification(formContext);
                        }, 5000);
                    } else {
                        const message = `Failed to advance phase from Draw to End Of Life.`;
                        setErrorFormNotification(formContext, message);
                    }
                });
            } else {
                const message = `Failed to advance phase from Order to Draw.`;
                setErrorFormNotification(formContext, message);
            }
        });
    } else if (isCurrentPhaseDraw(formContext)) {
        // From Draw to End Of Life
        formContext.data.process.moveNext(res => {
            if (res === "success") {
                const message = `Phase successfully advanced from Draw to End Of Life.`;
                setInfoFormNotification(formContext, message);

                setTimeout(() => {
                    clearInfoFormNotification(formContext);
                }, 5000);
            } else {
                const message = `Failed to advance phase from Draw to End Of Life.`;
                setErrorFormNotification(formContext, message);
            }
        });
    }
}