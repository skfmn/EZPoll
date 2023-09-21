function confirmSubmit(imsg, ihref) {
    var smsg = confirm(imsg);
    if (smsg == true) {
        window.location = ihref;
    } else {
        return false;
    }
}

function togglePass(arg1, arg2) {
    var x = document.getElementById(arg1);
    var y = document.getElementById(arg2);
    if (x.type === "password") {
        x.type = "text";
        y.className = "fa fa-eye shpwd";
    } else {
        x.type = "password";
        y.className = "fa fa-eye-slash shpwd";
    }
}

function validatePwd(form) {
    with (window.document.password) {
        if (cname.value == "") {
            alert('Please enter a name!');
            cname.focus();
            return false;
        }
        if (cpwd.value == "") {
            alert('Please enter a password!');
            cpwd.focus();
            return false;
        }
        if (cpwd2.value == "") {
            alert('Please enter the password again!');
            cpwd2.focus();
            return false;
        }
        if (cpwd.value != cpwd2.value) {
            alert('Passwords did not match\, please try again!');
            cpwd.focus();
            return false;
        }
        return true;
    }
}

function validateChangePwd(form) {
    with (window.document.password) {
        if (pwd.value == "") {
            alert('Please enter a password!');
            pwd.focus();
            return false;
        }
        if (pwd2.value == "") {
            alert('Please enter the password again!');
            pwd2.focus();
            return false;
        }
        if (pwd.value != pwd2.value) {
            alert('Passwords did not match\, please try again!');
            pwd.focus();
            return false;
        }
        return true;
    }
}

var pollOptionNum = 0, pollTabIndex, pollOptionNumMinus = 0;
function addPollOption() {
    if (pollOptionNum == 0) {
        for (var i = 0; i < document.reply.elements.length; i++)
            if (document.reply.elements[i].id.substr(0, 8) == "options-") {
                pollOptionNum++;
                pollTabIndex = document.reply.elements[i].tabIndex+1;
            }
    }
    pollOptionNum++
    pollOptionNumMinus = pollOptionNum - 1;
    setOuterHTML(document.getElementById("pollMoreOptions"), '<label for="options-' + pollOptionNumMinus + '" style="margin-bottom:-3px;">Answer ' + pollOptionNum + '</label><input type="text" name="options' + pollOptionNum + '" id="options-' + pollOptionNumMinus + '" tabindex="' + pollTabIndex + '" size="20" style="margin-bottom:10px;" value="" /><br /><span id="pollMoreOptions"></span>');
    //alert(pollOptionNum);
}

function setOuterHTML(element, toValue) {
    if (typeof (element.outerHTML) != 'undefined')
        element.outerHTML = toValue;
    else {
        var range = document.createRange();
        range.setStartBefore(element);
        element.parentNode.replaceChild(range.createContextualFragment(toValue), element);
    }
}