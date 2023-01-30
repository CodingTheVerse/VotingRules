def generatePreferences(values) -> dict:
    """
        generatePreferences() inputs a set of numerical values that the agents 
        have for the different alternatives and outputs a preference profile.
        Parameters:
            values : worksheet - worksheet corresponding to an xlsx file.
        Returns:
            pref : dict - a dictionary where the keys are the agents and the values are 
            lists that correspond to the preference orderings of those agents.
    """
    sheet=values
    pref={}
    agent=[]
    for row in sheet:
        list={}
        sortedList = []
        agentSel=[]
        dicts={}
        for cell in row:
            candidate=cell.column
            agent=cell.row
            score=cell.value
            list[candidate]=score
            sortedList = sorted(list.items(), key=lambda kv: (kv[1], kv[0]), reverse=True)
            agentSel = [i[0] for i in sortedList] 
        dicts[agent]=agentSel
        pref[agent] = dicts.get(agent)
    return (pref)
 
def dictatorship(preferenceProfile, agent) -> int:
    """
        dictatorship() determines the winner according to Dictatorship rule i.e. An agent 
        is selected, and the winner is the alternative that this agent ranks first. 
        Parameters:
            preferenceProfile : dict - preference profile represented by a dictionary
        Returns:
            dictWin : int - winner according to the Dictatorship rule
    """
    try:
        if agent not in preferenceProfile.keys():
            raise ValueError(agent, "This agent not valid")
        profile = []
        for key,value in preferenceProfile.items():
            if key == agent:
                profile = value
        dictWin = profile[0]

        #returns the first preference of the agent
        return dictWin
    except Exception as exp:
        print (exp)

def scoringRule(preferences, scoreVector, tieBreak) -> int :
    """
        scoringRule() determine a winner using the scores assigned to the alternatives 
        based on agent preference and the scoring vector. Have a tie-breaking option to 
        distinguish between alternatives with the same score.
        Parameters:
            preferences : dict - preference profile represented by a dictionary
            scoreVector : list - score vector of length 
            tieBreak : string - option for the tie-breaking among possible winners
        Returns:
            TSWinner : int - alternative with the highest total score
    """
    try:
        if len(scoreVector) != len(preferences[1]):
            raise ValueError("Incorrect input")

        alt = {}
        scoringPref = {}
        scoreVectorAsc = sorted(scoreVector, reverse = True)

        #assigns score to each alternative based on the scoring vector
        for key,value in preferences.items():
                scoringList = {}
                m = len(value)
                for b in range(m):                  
                    scoringList[preferences[key][b]] = scoreVectorAsc[b]    
                scoringPref[key] = scoringList      

        #calculate the totalscore for each alternative based on the scores assigned
        for key,value in scoringPref.items():
                for j in range(len(scoreVector)):
                    i = j+1
                    if i not in alt :
                        alt[i] =  value[i]
                    else:
                        alt[i] = alt[i] + value[i]
        maxScore = max(alt.values())
        maxFilter = [k for k, v in alt.items() if v == maxScore] 

        #calls the tiebreaker function in case of a tie
        if  len(maxFilter)  > 1:
            tieData = {}
            for t in range(len(maxFilter)):
                tieData[maxFilter[t]] = maxScore
            TSWinner = tiebreaker(tieBreak,tieData,preferences)
        else:
            TSWinner = (int(maxFilter[0]))
        return TSWinner
    except Exception as exp:
        print (exp)
        return False

def plurality (preferences, tieBreak) -> int:
    """
        plurality() determines the winner according to Plurality rule i.e. The winner is the 
        alternative that appears the most times in the first position of the agents' preference 
        orderings. Have a tie-breaking option to distinguish between alternatives with the same score.
        Parameters:
            preferences : dict - preference profile represented by a dictionary
            tieBreak : string - option for the tie-breaking among possible winners
        Returns:
            PLWinner : int - alternative that appears the most times in the first position 
            of the agents' preference
    """
    alt = {}
    for key,value in preferences.items():
                for j in range(len(value)):
                    if j == 0 :
                        if value[0] not in alt :
                            alt[value[0]] =  1
                        else:
                            alt[value[0]] = alt[value[0]] + 1 
    maxOcc =  max(alt.values())
    maxFilter = [k for k, v in alt.items() if v == maxOcc] 

    #calls the tiebreaker function in case of a tie
    if len(maxFilter)  > 1:
        tieData = {}
        for t in range(len(maxFilter)):
            tieData[maxFilter[t]] = maxOcc
        PLWinner = tiebreaker(tieBreak,tieData,preferences)
    else:
        PLWinner = (int(maxFilter[0]))
    return PLWinner

def veto (preferences, tieBreak) -> int:
    """
        veto() determines the winner according to Veto rule i.e. Every agent assigns 0 points 
        to the alternative that they rank in the last place of their preference orderings, and 
        1 point to every other alternative. Have a tie-breaking option to distinguish between 
        alternatives with the same score.
        Parameters:
            preferences : dict - preference profile represented by a dictionary
            tieBreak : string - option for the tie-breaking among possible winners
        Returns:
            vetoWinner : int - The winner is the alternative with the most number of points.
    """
    alt = {}
    vetoPref = {}
    for key,value in preferences.items():
                vetoList = {}
                m = len(value)
                for v in range(m):                  
                    if  v == m-1 :
                        vetoList[preferences[key][v]] = 0
                    else:
                        vetoList[preferences[key][v]] = 1
                vetoPref[key] = vetoList 

    #calculate the totalscore for each alternative based on the scores assigned (i.e. 0 or 1)                                      
    for key,value in vetoPref.items():
                for j in range(len(value)):
                    i = j + 1
                    if i not in alt :
                        alt[i] =  value[i]
                    else:
                        alt[i] = alt[i] + value[i]  
    maxPoints =  max(alt.values())
    maxFilter = [k for k, v in alt.items() if v == maxPoints]

    #calls the tiebreaker function in case of a tie
    if len(maxFilter)  > 1:
        tieData = {}
        for t in range(len(maxFilter)):
            tieData[maxFilter[t]] = maxPoints
        vetoWinner = tiebreaker(tieBreak,tieData,preferences)
    else:
        vetoWinner = (int(maxFilter[0]))
    return vetoWinner

def borda (preferences, tieBreak) -> int:
    """
        borda() determines the winner according to borda rule i.e. the alternative ranked at 
        position j receives a score of m-j. The winner is the alternative with the highest score. 
        Have a tie-breaking option to distinguish between alternatives with the same score.
        Parameters:
            preferences : dict - preference profile represented by a dictionary
            tieBreak : string - option for the tie-breaking among possible winners
        Returns:
            bordaWinner : int - The winner is the alternative with the highest score.
    """
    alt = {}
    bordaPref = {}
    for key,value in preferences.items():
                bordaList = {}
                m = len(value)
                for b in range(m):                  
                    bordaList[preferences[key][b]] = (m - 1) - b
                bordaPref[key] = bordaList

    #calculate the totalscore for each alternative based on the scores assigned                                                                    
    for key,value in bordaPref.items():
                for j in range(len(value)):
                    i = j + 1
                    if i not in alt :
                        alt[i] =  value[i]
                    else:
                        alt[i] = alt[i] + value[i]  
    maxPoints =  max(alt.values())
    maxFilter = [k for k, v in alt.items() if v == maxPoints]

    #calls the tiebreaker function in case of a tie
    if len(maxFilter)  > 1:
        tieData = {}
        for t in range(len(maxFilter)):
            tieData[maxFilter[t]] = maxPoints
        bordaWinner = tiebreaker(tieBreak,tieData,preferences)
    else:
        bordaWinner = (int(maxFilter[0]))
    return bordaWinner

def harmonic (preferences, tieBreak) -> int:
    """
        harmonic() determines the winner according to harmonic rule i.e. the alternative ranked at 
        position j receives a score of 1/j. The winner is the alternative with the highest score. 
        Have a tie-breaking option to distinguish between alternatives with the same score.
        Parameters:
            preferences : dict - preference profile represented by a dictionary
            tieBreak : string - option for the tie-breaking among possible winners
        Returns:
            harmonicWinner : int - The winner is the alternative with the highest score.
    """
    alt = {}
    harPref = {}
    for key,value in preferences.items():
                harList = {}
                m = len(value)
                for b in range(m):                  
                    harList[preferences[key][b]] = 1 / (b+1)
                harPref[key] = harList
                
    #calculate the totalscore for each alternative based on the scores assigned                                                                    
    for key,value in harPref.items():
                for j in range(len(value)):
                    i = j + 1
                    if i not in alt :
                        alt[i] =  value[i]
                    else:
                        alt[i] = alt[i] + value[i]  
    maxPoints =  max(alt.values())
    maxFilter = [k for k, v in alt.items() if v == maxPoints]

    #calls the tiebreaker function in case of a tie
    if len(maxFilter)  > 1:
        tieData = {}
        for t in range(len(maxFilter)):
            tieData[maxFilter[t]] = maxPoints
        harmonicWinner = tiebreaker(tieBreak,tieData,preferences)
    else:
        harmonicWinner = (int(maxFilter[0]))
    return harmonicWinner

def STV (preferences, tieBreak) -> int:
    """
        STV() determines the winner according to Single Transferable Vote rule i.e. the alternatives 
        that appear the least frequently in the first position of agents' rankings are removed, and 
        the process is repeated. The last set is the set of possible winners. Have a tie-breaking 
        option to distinguish between alternatives with the same score.
        Parameters:
            preferences : dict - preference profile represented by a dictionary
            tieBreak : string - option for the tie-breaking among possible winners
        Returns:
            STVWinner : int - The winner according to Single Transferable Vote rule .
    """
    STVPref = preferences
    STVWinner = ''
    for i in range(len(STVPref[1])) :
        # checks if the number of alternative left is greater than 1. If true , removes the least frequent alternative on first place.
        # If false, goes to elif and returns the alternative as the winner. 
        if len(STVPref[1]) > 1 :
            m = len(STVPref[1])
            alt = {}
            leastFreq={}
            for key,value in STVPref.items():
                        for j in range(len(value)):
                            if j == 0 :
                                if value[0] not in alt :
                                    alt[value[0]] =  1
                                else:
                                    alt[value[0]] = alt[value[0]] + 1 
            for key,value in STVPref.items():
                for k in range(m):
                        if value[k] in alt.keys():
                            leastFreq[value[k]] = alt[value[k]]
                        else:
                            leastFreq[value[k]] = 0
            minOcc =  min(leastFreq.values())
            minFilter = [k for k, v in leastFreq.items() if v == minOcc] 

            #calls the tiebreaker function in case of a tie
            if len(minFilter)  > 1:
                tieData = {}
                for t in range(len(minFilter)):
                    tieData[minFilter[t]] = minOcc
                STVWinner = tiebreaker(tieBreak,tieData,preferences)
            else:
                STVWinner = int(minFilter[0])
            
            #removes the 'least frequent alternative on first place' from STVPref. and continue. 
            for r in range(len(STVPref)):
                s = r + 1
                STVPref[s].remove(STVWinner)               
        elif len(STVPref[1]) == 1 :
            STVWinner = int(STVPref[1][0])
            return STVWinner
    
def rangeVoting (values, tieBreak) -> int:
    """
        rangeVoting() The function should return the alternative that has the maximum 
        sum of valuations, i.e., the maximum sum of numerical values in the xlsx file, 
        using the tie-breaking option to distinguish between possible winners.
        Parameters:
            values : worksheet - worksheet corresponding to an xlsx file.
            tieBreak : string - option for the tie-breaking among possible winners
        Returns:
            RVWinner : int - alternative that has the maximum sum of valuations 
    """
    sheet=values
    maxSum={}
    for row in sheet:
        for cell in row:
            candidate=cell.column
            agent=cell.row
            score=cell.value
            if candidate not in maxSum :
                maxSum[candidate] =  score
            else:
                maxSum[candidate] = maxSum[candidate] + score 
    winScore =  max(maxSum.values())
    maxScore = [k for k, v in maxSum.items() if v == winScore] 

    #calls the tiebreaker function in case of a tie
    if len(maxScore)  > 1:
            tieData = {}
            data = generatePreferences(values)
            for t in range(len(maxScore)):
                tieData[maxScore[t]] = winScore
            RVWinner = tiebreaker(tieBreak,tieData,data)
    else:
            RVWinner = (int(maxScore[0]))
    return RVWinner

def tiebreaker(tieBreak,tieData,preferences):
    """
        tiebreaker() The function distinguish between possible winners.
        Parameters:
            tieBreak : string - option for the tie-breaking among possible winners.
            tieData : dict - Possible winners with a tie
            preferences : dict - preference profile represented by a dictionary 
        Returns:
            tieWin : int - alternative that wins based on the tieBreak input 
    """
    agentList=[]
    if tieBreak == 'max':
        tieWin =  max(tieData.keys())
    elif tieBreak =='min':
        tieWin =  min(tieData.keys())
    elif tieBreak in preferences.keys():
         agentChoice = [v for k, v in preferences.items() if k == tieBreak][0]
         for i in tieData :
             agentList.append(agentChoice.index(i))
         tieWin = agentChoice[min(agentList)]
    return tieWin


