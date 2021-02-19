The simulation takes in two types of inputs - a file containing all the details for the ecosystems, and a set of files containing data for specific users.
Sample files are included in this folder.

## General Information

* The order that ecosystems are listed should remain uniform throughout all the sheets
* The simulation can only simulate up to the time periods there is data for 
* The names of the files should remain the same
* The final results will be exported into a excel file titled "export" 


## ReputationDetails File

The ReputationDetails file has three sheets, Congruency, Details, and Initial Reputations

### Congruency
* This sheet controls how different ecosystems are weighted by each other
* The rows of the table in the details sheet represent the weighting/congruenncy of a specific reputation
* The columns indicate which ecosystem receives that rating 
* The ratings should all add up to 1 for the simulation to work properly
  
### Details
* This sheet is concered with reputation design
* The high cutoff is the minimum reputation needed to be considered a highly rated user
* The high vaue is the weight of a high users vote
* The low value is the weight of a low users vote
* Update periods are when the ecosystem updates reputation
  * Ex, an update period of 3 means a ecosystem will update its reputation for any time period that is a multiple of 3
* Legacy will be added next iteration
    
### Initial Reputations
* This sheet contains the initial reputations of all users in the simulation
* If a user does not participate in a ecosystem, leave the score blank

## Reputation Sheets

* Each sheet contains the amount of votes a user gets
* A vote is represented by a 1, and not voting is represented by a 0
  * If a user does not participate in the specific ecosystem, their row should be left blank 
* The sheet names represent the specific ecosystem the votes are for 
