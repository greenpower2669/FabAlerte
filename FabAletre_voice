ObjC.import('Cocoa');
ObjC.import('stdlib');

// Définir l'application de la voix
var speechSynth = $.NSSpeechSynthesizer.alloc.init;

// Fonction pour parler
function speak(text) {
    speechSynth.startSpeakingString(text);
}

// Fonction principale
function main() {
    var now = new Date();
    var alertTime = new Date(now.getTime() + 6 * 60000); // Mettre l'alerte pour dans 6 minutes
    var userTimeInput = prompt("Veuillez entrer l'heure de l'alerte (format HH:MM, par exemple " + alertTime.getHours() + ":" + alertTime.getMinutes() + ")", "Entrée d'alerte");

    if (!userTimeInput) {
        speak("Aucune heure entrée, le programme va s'arrêter.");
        return;
    }

    var alertTime = new Date();
    alertTime.setHours(userTimeInput.split(":")[0]);
    alertTime.setMinutes(userTimeInput.split(":")[1]);

    if (isNaN(alertTime.getTime()) || new Date() >= alertTime) {
        speak("Heure invalide ou déjà passée.");
        return;
    }

    var interval = setInterval(function() {
        var now = new Date();
        if (now >= alertTime) {
            clearInterval(interval);
            speak("L'heure de l'alerte est passée. Veuillez entrer une nouvelle heure pour la prochaine alerte.");
            main();  // Relancer la demande d'heure
        } else {
            var minutesRemaining = Math.round((alertTime - now) / 60000);
            if (minutesRemaining <= 1) {
                speak("Ceci est un rappel! Il reste maintenant moins d'une minute avant l'heure de votre alerte.");
            }
        }
    }, 30000); // Vérifie toutes les 30 secondes
}

main(); // Appel de la fonction principale
