/* Copyright (c) Microsoft Corporation. All rights reserved.
   Licensed under the MIT license. See LICENSE in the project root for license information. */

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("run").onclick = run;
    }
  });
  
  export async function run() {
    return Word.run(async (context) => {
      const body = context.document.body;
  
      // Vérifier si le document contient déjà du texte
      body.load("text");
      await context.sync();
  
      if (body.text.trim().length > 0) {
        const confirmation = confirm("La page contient déjà du texte. Voulez-vous effacer le contenu ?");
        if (!confirmation) {
          return; // On arrête si l'utilisateur ne veut pas effacer
        }
        body.clear();
      }
  
      // Ajout de texte répété pour générer 1 à 2 pages
      for (let i = 0; i < 10; i++) {
        body.insertParagraph(
          "Consigne : Insérez une image à cet emplacement, alignez-la à droite et assurez-vous que le texte s’adapte correctement autour de l’image.",
          Word.InsertLocation.end
        );
      }
  
      // Insérer une image exemple (base64 1x1 pixel si rien de dispo)
      // Remplace "<BASE64_IMAGE>" par une image encodée en Base64 ou un logo de ton choix
      const placeholderImage =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIW2NkYGBgAAAABAABJzQnCgAAAABJRU5ErkJggg==";
  
      const image = body.insertInlinePictureFromBase64(
        placeholderImage,
        Word.InsertLocation.end
      );
  
      // Mise en forme : alignement à droite + habillage
      image.floatingFormat.horizontalAlignment = "Right";
      image.floatingFormat.wrapFormat.type = "Square"; // texte autour de l'image
  
      await context.sync();
  
      // Message de confirmation pour l'utilisateur
      console.log("Document généré avec texte + image alignée à droite.");
      alert("Le document a été généré : texte + image alignée à droite avec habillage.");
    });
  }
  