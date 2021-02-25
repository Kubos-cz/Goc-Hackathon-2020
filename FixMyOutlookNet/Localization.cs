using Microsoft.PowerShell.Commands;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FixMyOutlookNet
{

    static class Localization
    {
        private static CultureInfo Language { get; set; } = CultureInfo.InstalledUICulture;

        public static string GetUIText(int Index)
        {
            switch (Language.ThreeLetterISOLanguageName) 
            {
                case "nld":
                    return Dutch[Index];
                case "ita":
                    return Italian[Index];
                case "deu":
                       return German[Index];
                case "fra":
                    return French[Index];
                case "eng":
                default:
                    return English[Index];
            }
        }

        static Dictionary<int, string> English = new Dictionary<int, string>()
        {
            {1,"|| OUTLOOK CONFIGURATION FIX ||" },
            {2,"Outlook has to be closed to proceed, press any key to continue" },
            {3,"Closing outlook..." },
            {4,"Detecting office instalation..." },
            {5,"Detected version:" },
            {6,"Cleaning auto-generated profiles..." },
            {7,"Creating profile..." },
            {8,"Preparing outlook configuration..." },
            {9,"Outlook configuration has finished you may now start outlook, press any key to continue." }
        };

        static Dictionary<int, string> Italian = new Dictionary<int, string>()
        {
            {1,"|| RIPARA IL MIO OUTLOOK ||" },
            {2,"Outlook deve essere chiuso, premere un tasto per continuare" },
            {3,"Chiusura Outlook in corso..." },
            {4,"Ricerca versione di Office corrente…" },
            {5,"Versione di Office corrente:" },
            {6,"Pulizia profili auto–generati..." },
            {7,"Creazione profilo in corso..." },
            {8,"Configurazione Outlook in corso..." },
            {9,"La configurazione di Outlook è terminata, adesso può avviare Outlook, premere un tasto per continuare." }
        };

        static Dictionary<int, string> German = new Dictionary<int, string>()
        {
            {1,"|| OUTLOOK REPARIEREN ||" },
            {2,"Outlook muss geschlossen sein, um fortzufahren, drücken Sie eine beliebige Taste, um fortzufahren" },
            {3,"Outlook wird geschlossen..." },
            {4,"Erkennung der Office-Installation..." },
            {5,"Erkannte Version:" },
            {6,"Automatisch erzeugte Profile bereinigen..." },
            {7,"Profil wird erstellt..." },
            {8,"Einrichten der Outlook-Konfiguration..." },
            {9,"Outlook-Konfiguration ist abgeschlossen, Sie können nun Outlook starten, drücken Sie eine beliebige Taste, um fortzufahren." }
        };

        static Dictionary<int, string> French = new Dictionary<int, string>()
        {
            {1,"|| RÉPARER MON OUTLOOK ||" },
            {2,"Outlook doit être fermé pour continuer, appuyez sur n'importe quelle touche pour continuer" },
            {3,"Fermeture de Outlook..." },
            {4,"Détection de l'installation office..." },
            {5,"Version détectée:" },
            {6,"Nettoyage des profils générés automatiquement..." },
            {7,"Créer un profil..." },
            {8,"Configurer la configuration de outlook..." },
            {9,"La configuration d'Outlook est terminée, vous pouvez maintenant démarrer Outlook, appuyez sur n'importe quelle touche pour continuer." }
        };

        static Dictionary<int, string> Dutch = new Dictionary<int, string>()
        {
            {1,"|| OUTLOOK CONFIGURATIE REPARATIE || " },
            {2,"Outlook moet gesloten zijn om verder te gaan, druk op een willekeurige toets om verder te gaan" },
            {3,"Outlook sluiten..." },
            {4,"Detecteren van Office installatie..." },
            {5,"Gedetecteerde versie:" },
            {6,"Opschonen automatisch gegenereerde profielen..." },
            {7,"Aanmaken profiel..." },
            {8,"Opzetten van outlook configuratie..." },
            {9,"Outlook configuratie is klaar u kunt nu outlook starten, druk op een willekeurige toets om verder te gaan." }
        };


    }
}
