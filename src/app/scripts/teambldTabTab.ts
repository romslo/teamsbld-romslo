import { TeamsTheme } from './theme';

/**
 * Implementation of the teambldTab content page
 */
export class teambldTabTab {
    /**
     * Constructor for teambldTab that initializes the Microsoft Teams script and themes management
     */
    constructor() {
        microsoftTeams.initialize();
        TeamsTheme.fix();
    }
    /**
     * Method to invoke on page to start processing
     * Add your custom implementation here
     */
    public doStuff() {
        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            let element = document.getElementById('app');
            if (element) {
                element.innerHTML = `The value is: ${context.entityId}`;
            }
        });
    }
}