
import services_handler as service_handler

MENU_OPTIONS = {
    1: 'Generer la feuille de temps zoom', 
    2: 'Generer les attestations de présences', 
    3: 'En cours....',
    4: 'Quitter' 
}

def run_option(option):
    """
    Execute la fonctionnalité selon le choix de l'utilisateur
    """
    if option not in range(1, len(MENU_OPTIONS) + 1):
        raise ValueError('Merci de saisir une option valide')

    print(f'[{option}] - {MENU_OPTIONS[option].upper()}')
    match option:
        case 1:
            service_handler.generate_timesheet_zoom()
        case 2:
            service_handler.generate_attendance_certificates()
        case 4:
            exit()

def show_options():
    """
    Affiche les options de l'application
    """
    for key in range(1, len(MENU_OPTIONS) + 1):
        print(f'[{key}] - {MENU_OPTIONS[key]}')

def start():
    while True:
        show_options()
        try:
            run_option(int(input('Merci de saisir une option valide: ')))
        except KeyboardInterrupt:
            print('\nAurevoir...')
            exit()
        except ValueError:
            pass

start()