let sellsWH = "https://discord.com/api/webhooks/1085283582039293992/odf5TlDH4cLCYsLZoPU-YaPSIbuHI4Cx9t7joLrrn7J-OdwcR8iubq9rIoGT8Le6xzce";
let primeWH = "https://discord.com/api/webhooks/1086335597876092989/RTc7ndo-THdB13DD8125fhMvtHwVmbFMFi549yFAr01qZnxukxxFrKuo1IrJl7KRsaE0";
let presenceWH = "https://discord.com/api/webhooks/1123597989471068240/JiclZ7vuWexkd7gOIEsKyGY68InybnCX9ov1dMJcKl_lrCUdEYg6gp9BFZdyvO4h6t5h";

function addSell() {
  
  let sheet = SpreadsheetApp.getActive();
  let sells = SpreadsheetApp.getActive().getSheetByName("Ventes");
  let parameters = sheet.getRange("C4:C11");

  var id = parameters.getCell(1, 1).getValue();
  var author = parameters.getCell(2, 1).getValue();
  var ticket = parameters.getCell(3, 1).getValue();
  var type = parameters.getCell(5, 1).getValue();
  var price = parameters.getCell(6, 1).getValue();
  var interior = parameters.getCell(7, 1).getValue();
  var garage = parameters.getCell(8, 1).getValue();
  let date = Utilities.formatDate(new Date(), "GMT+2", "dd/MM/yy");

  if(id != "" && author != "" && ticket != "" && type != "" && price != "" && interior != "" && garage != "") {
    var response = SpreadsheetApp.getUi().alert("✅ Votre vente a été ajoutée avec succès", "Voici le récapitulatif de la vente que vous avez ajoutée :\n- Compte client: " + id + "\n- Agent Immobilier: " + author + "\n- Ticket: " + ticket + "\n- Type de propriété: " + type + "\n- Prix: " + price + "$\n- Intérieur: " + interior + "\n- Garage: " + garage + "\n\n❔ Souhaitez-vous récupérer votre prime ? (" + getSellerPrime(sells, author) + "$)\nLe bouton \"Annuler\" supprimera votre vente.", SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL);
    switch(response) {
      case SpreadsheetApp.getUi().Button.YES:
        SpreadsheetApp.getUi().alert("💰 Votre demande a été transmise à la direction", "Merci de patienter en attendant que l'administration réponde à votre demande. Vous serez recontacté sur l'intranet\n\nMontant de la prime: " + getSellerPrime(sells, author) + "$", SpreadsheetApp.getUi().ButtonSet.OK);
        sendDiscordMessage(primeWH, "<@&1084923439774703706>\n\n**" + author + "** a demandé à __récupérer sa prime__.\nMontant : `" + getSellerPrime(sells, author) + "$`");
        break;

      case SpreadsheetApp.getUi().Button.CANCEL:
        SpreadsheetApp.getUi().alert("✅ Votre vente a été annulée", "La vente a été annulée et ne sera pas publiée ni sur le document de la comptabilité ni sur l'intranet.", SpreadsheetApp.getUi().ButtonSet.OK);
        reset(parameters, 1);
        clear();
        return;

      default:
        break;
    }

    let data = getSellerRange(sells, author);
    var x = 1;
    while(!data.getCell(x, 1).isBlank()) {
      x += 1;
    }
    data.getCell(x, 1).setValue(date);
    data.getCell(x, 2).setValue(id);
    data.getCell(x, 3).setValue(ticket);
    data.getCell(x, 4).setValue(type);
    data.getCell(x, 5).setValue(interior);
    data.getCell(x, 6).setValue(garage);
    data.getCell(x, 7).setValue(price);

    reset(parameters, 1);
    clear(1);
    pushToStats(price);
    sendDiscordMessage(sellsWH, getSellerDiscord(author) + " a réalisé une vente ||" + ticket + "||\n> Type de propriété: **" + type + "**\n> Prix: `" + price + "$`");

  }else {
    SpreadsheetApp.getUi().alert("⚠️ Des informations sont manquantes", "Vous devez remplir toutes les cases afin d'ajouter une vente.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function addAccount() {
  let sheet = SpreadsheetApp.getActive();
  let parameters = sheet.getRange("F4:F9");
  let accounts = SpreadsheetApp.getActive().getSheetByName("Clients");

  var identity = parameters.getCell(1, 1).getValue();
  var phone = parameters.getCell(2, 1).getValue();
  var discord = parameters.getCell(3, 1).getValue();
  var public = parameters.getCell(4, 1).getValue();
  var ceo = parameters.getCell(5, 1).getValue();
  var dynasty = parameters.getCell(6, 1).getValue();

  if(identity != "" && phone != "" && discord != "") {
    var id = Math.floor(Math.random()*1000000000);
    let data = accounts.getRange("B4:H1000");
    var x = 1;
    while(!data.getCell(x, 1).isBlank()) {
      x += 1;
    }

    data.getCell(x, 1).setValue(id);
    data.getCell(x, 2).setValue(identity);
    data.getCell(x, 3).setValue(phone);
    data.getCell(x, 4).setValue(discord);
    data.getCell(x, 5).setValue((public ? "Oui" : "Non"));
    data.getCell(x, 6).setValue((ceo ? "Oui" : "Non"));
    data.getCell(x, 7).setValue((dynasty ? "Oui" : "Non"));

    reset(parameters, 2);
    clear(2);

  SpreadsheetApp.getUi().alert("✅ Votre compte client a été ajouté", "Voici le récapitulatif du compte client que vous avez ajouté :\n- Identité: " + identity + "\n- Numéro de téléphone: " + phone + "\n- Discord: " + discord + "\n- Service Public ? " + (public ? "Oui" : "Non") + "\n- Patron d'entreprise ? " + (ceo ? "Oui" : "Non") + "\n- Employé D8 ? " + (dynasty ? "Oui" : "Non") + "\n\n- Numéro de compte client : " + id, SpreadsheetApp.getUi().ButtonSet.OK);

  }else {
    SpreadsheetApp.getUi().alert("⚠️ Des informations sont manquantes", "Vous devez remplir toutes les cases afin d'ajouter un compte client.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function addPresence1() {
  addPresence(SpreadsheetApp.getActive().getSheetByName("Dashboard").getRange("B15:C17"));
}

function addPresence2() {
  addPresence(SpreadsheetApp.getActive().getSheetByName("Dashboard").getRange("E15:F17"));
}

function addPresence3() {
  addPresence(SpreadsheetApp.getActive().getSheetByName("Dashboard").getRange("H15:I17"));
}

function addPresence(range) {
  let identity = range.getCell(1, 2).getValue();
  let start = range.getCell(2, 2).getValue();
  let stop = range.getCell(3, 2).getValue();
  if(identity != "" && start != "" && stop != "") {
    sendDiscordMessage(presenceWH, getSellerDiscord(identity) + " a signalé être présent à l'agence immobilière\n> Début: `" + start + "`\n> Fin: `" + stop + "`");
    reset(range, 3);
    clear(3);
    SpreadsheetApp.getUi().alert("✅ Votre présence a été enregistrée", "Votre présence à l'agence immobilière a été enregistrée et signalée à la Direction. Votre présence ne sera pas affichée dans le document de la comptabilité avant que la Direction ait vérifié que vous étiez bien à l'agence immobilière à ce moment là.", SpreadsheetApp.getUi().ButtonSet.OK);
  }else {
    SpreadsheetApp.getUi().alert("⚠️ Des informations sont manquantes", "Vous devez remplir toutes les cases afin d'ajouter une présence.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function registerPresence() {
  let sells = SpreadsheetApp.getActive().getSheetByName("Ventes");
  let direction = SpreadsheetApp.getActive().getSheetByName("Direction");
  let parameters = direction.getRange("Q38:R46");

  let date = Utilities.formatDate(new Date(), "GMT+2", "dd/MM/yy");
  let author = parameters.getCell(3*0+1, 1).getValue();
  let salary = parameters.getCell(3*1+1, 1).getValue();
  let duration = parameters.getCell(3*2+1, 1).getValue();

  if (author != "" && salary != "" && duration != "") {
    let data = getSellerRange(sells, author);
    var x = 1;
    while(!data.getCell(x, 1).isBlank()) {
      x += 1;
    }
    data.getCell(x, 1).setValue(date);
    data.getCell(x, 2).setValue("-----");
    data.getCell(x, 3).setValue("-----");
    data.getCell(x, 4).setValue("$" + salary + " x " + duration + "mn");
    data.getCell(x, 5).setValue("-----");
    data.getCell(x, 6).setValue("-----");
    data.getCell(x, 7).setValue(salary*duration);

    reset(parameters, 4);

    SpreadsheetApp.getUi().alert("✅ La présence a été enregistrée", "Voici le récapitulatif de la présence que vous avez ajouté :\n- Employé: " + author + "\n- Salaire: $" + salary + "/min\n- Durée: " + duration + " minute(s)\n- Prime: $" + salary*duration, SpreadsheetApp.getUi().ButtonSet.OK);
    sendDiscordMessage(sellsWH, "La présence de " + getSellerDiscord(author) + " a été validée\n> Durée: `" + duration + " minute(s)`\n> Prime: `" + salary*duration + "$`");
  }else {
    SpreadsheetApp.getUi().alert("⚠️ Des informations sont manquantes", "Vous devez remplir toutes les cases afin d'ajouter une présence.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function pushToStats(cost) {
  let sheet = SpreadsheetApp.getActive().getSheetByName("Direction");
  let data = sheet.getRange("J18:M1000");
  let date = Utilities.formatDate(new Date(), "GMT+2", "dd/MM/yy");

  var x = 1;
  while(!data.getCell(x, 1).isBlank() && data.getCell(x, 1).getValue() != date) {
    x += 1;
  }
  if(data.getCell(x, 1).getValue() != date) {
    data.getCell(x, 1).setValue(date);
    data.getCell(x, 2).setValue(1);
    data.getCell(x, 3).setValue(cost);
    data.getCell(x, 4).setValue("=L" + (17+x) + "*0,5")
  }else {
    data.getCell(x, 2).setValue(data.getCell(x,2).getValue()+1);
    data.getCell(x, 3).setValue(data.getCell(x,3).getValue()+cost);
  }
}

function getSellerRange(sheet, name) {
  switch(name) {
    case "T. Clark":
      return sheet.getRange("B6:H1000");
    
    case "M. Hendrix":
      return sheet.getRange("J6:P1000");
      
    case "A. Antranik":
      return sheet.getRange("R6:X1000");

    case "D. Walter":
      return sheet.getRange("Z6:AF1000");

  }
}

function getSellerPrime(sheet, name) {
  switch(name) {
    case "T. Clark":
      return sheet.getRange("H4").getValue();
    
    case "M. Hendrix":
      return sheet.getRange("P4").getValue();
      
    case "A. Antranik":
      return sheet.getRange("X4").getValue();
      
    case "D. Walter":
      return sheet.getRange("AF4").getValue();

  }
}

function getSellerDiscord(name) {
  switch(name) {
    case "T. Clark":
      return "<@346352234914643981>";

    case "M. Hendrix":
      return "<@481442129177083912>";

    case "A. Antranik":
      return "<@573631432426258443>";

    case "G. Menfain":
      return "<@1097178680318496901>";

  }
}

function reset(range, value) {
  if(value == 1) { // Ajouter une vente
    range.getCell(1, 1).setValue("");
    range.getCell(2, 1).setValue("");
    range.getCell(3, 1).setValue("");
    range.getCell(5, 1).setValue("");
    range.getCell(6, 1).setValue("");
    range.getCell(7, 1).setValue("");
    range.getCell(8, 1).setValue("");
  }else if(value == 2) { // Ajouter un client
    range.getCell(1, 1).setValue("");
    range.getCell(2, 1).setValue("");
    range.getCell(3, 1).setValue("");
    range.getCell(4, 1).setValue("");
    range.getCell(5, 1).setValue("");
    range.getCell(6, 1).setValue("");
  }else if(value == 3) { // Présence Agence
    range.getCell(1, 2).setValue("");
    range.getCell(2, 2).setValue("");
    range.getCell(3, 2).setValue("");
  }else if(value == 4) { // Register Présence
    range.getCell(3*0+1, 1).setValue("");
    range.getCell(3*1+1, 1).setValue("");
    range.getCell(3*2+1, 1).setValue("");
  }
}

function clear(value) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var cells = [];
  if(value == 1) {
    cells = [sheet.getRange("C4:C6"), sheet.getRange("C8:C11")];
  }else if(value == 2) {
    cells = [sheet.getRange("F4:F9")];
  }else if(value == 3) {
    cells = [sheet.getRange("C15:C17"), sheet.getRange("F15:F17"), sheet.getRange("I15:I17")];
  }
  cells.forEach(range => {
    range.clearFormat();
    range.setBackgroundRGB(140, 207, 172);
    range.setFontColor("white")
    range.setFontSize(14);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
  });
}

function sendDiscordMessage(url, body) {
  var payload = JSON.stringify({content: body});
  var params = {
    headers: {"Content-Type": "application/json"},
    method: "POST",
    payload: payload,
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(url, params);
}
