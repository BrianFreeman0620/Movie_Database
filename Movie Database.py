# Imports the needed code to open excel files, create graphics windows, and
# quickly run math equations
import pandas as pd
from graphics import *
from math import *

# Finds the distance between two objects given two lists of numbers
def minkowskiDist(v1,v2,p):
    sumOfV = 0
    for value in range(len(v1)):
        try:
            sumOfV += (abs(float(v1[value]) - float(v2[value]))) ** p
        except:
            pass
    minkowskiValue = sumOfV ** (1/p)
    return minkowskiValue

# Using the minkowski distance, finds the closest [number] of movies
def nearestMovies(new, feature, number):
    similarMovies = []
    for i in range(number):
        minimum = "none"
        for vector in feature:
            if new == feature[vector]:
                pass
            elif vector in similarMovies:
                pass
            else:
                if minimum == "none":
                    minimum = minkowskiDist(new, feature[vector], 2)
                    neighbor = vector
                elif minimum > minkowskiDist(new, feature[vector], 2):
                    minimum = minkowskiDist(new, feature[vector], 2)
                    neighbor = vector
        similarMovies.append(neighbor)
        if i == 0:
            nearestMovie = neighbor
    return similarMovies, nearestMovie

# Checks if a button is clicked in the graphics window
def clicked(button,p):
    c=button.getCenter()
    r=button.getRadius()
    return sqrt((p.getX()-c.getX())**2+(p.getY()-c.getY())**2) < r

def main():
    # Imports the data from the excel file Move Data.xlsx
    infile = pd.read_excel('Movie Data.xlsx', 'Sheet1')
    tryMovie = "yes"
    
    # Creates lists for all values in the list by type
    movieList = []
    budgetList = []
    criticList = []
    audienceList = []
    salesList = []
    
    # Adds the data to the five lists
    for movie in infile["Name of Movie"]:
        movieList.append(movie)
    for budget in infile["Budget (Millions)"]:
        budgetList.append(budget)
    for critic in infile["Critic Score"]:
        criticList.append(critic)
    for audience in infile["Audience Score"]:
        audienceList.append(audience)
    for sales in infile["Box Office (Millions)"]:
        salesList.append(sales)
        
    movieDict = {}
    movieListCount = 0
    
    # Adds the data into the dictionary movieDict to easily be able to refer 
    # to later
    for movie in movieList:
        movieDict[movie] = [budgetList[movieListCount], 
                  criticList[movieListCount], audienceList[movieListCount], 
                  salesList[movieListCount]]
        movieListCount += 1
    
    # Creates the graphics window and text explaining what the program does
    win = GraphWin("Brian Freeman Final Project", 1000, 500)
    words = Text(Point(500,100),"""Using the data from hundreds of movies, this program will be able to predict the box office in millions of dollars \n
             of a hypothetical movie given its budget, critic scores, and audience scores. \n
             Warning! Budgets above 200 million dollars may be less accurate due to the limited of real movies \n
             above that amount \n
             Also, the run button may need to be clicked multiple times to work""")
    words.setSize(8)
    words.draw(win)
    
    # Asks the user for a hypothetical budget for a movie
    inputBudgetBox = Entry(Point(600,300),10)
    inputBudgetBox.setFill("white")
    inputBudgetBox.setTextColor("green")
    inputBudgetBox.draw(win)
    inputBudgetText = Text(Point(300,300), "In millions of dollars, what is the budget of the movie?")
    inputBudgetText.draw(win)
    
    # Asks the user for a hypothetical critic score for a movie
    inputCriticBox = Entry(Point(600,350),10)
    inputCriticBox.setFill("white")
    inputCriticBox.setTextColor("red")
    inputCriticBox.draw(win)
    inputCriticText = Text(Point(300,350), "From 0-100, what is the critic score for the movie?")
    inputCriticText.draw(win)
    
    # Asks the user for a hypothetical audience score for a movie
    inputAudienceBox = Entry(Point(600,400),10)
    inputAudienceBox.setFill("white")
    inputAudienceBox.setTextColor("blue")
    inputAudienceBox.draw(win)
    inputAudienceText = Text(Point(300,400), "From 0-100, what is the audience score for the movie?")
    inputAudienceText.draw(win)
    
    # Creates the button that compares inserted data to the dictionary's data
    clickCircle = Circle(Point(600,450),25)
    clickCircle.setFill("green")
    clickCircle.draw(win)
    clickCircleText = Text(Point(600,450),"Run")
    clickCircleText.draw(win)
    
    # Creates a film reel and draws it in the winow
    movieBase = Circle(Point(150,150), 75)
    movieBase.setFill("Black")
    movieBase.draw(win)
    
    movieFilm = Polygon([Point(225,150), Point(200,175), Point(250,225), Point(275,200)])
    movieFilm.setFill("Black")
    movieFilm.draw(win)
    
    movieHole1 = Circle(Point(200,150), 10)
    movieHole1.setFill("white")
    movieHole1.draw(win)
    
    movieHole2 = Circle(Point(150,200), 10)
    movieHole2.setFill("white")
    movieHole2.draw(win)
    
    movieHole3 = Circle(Point(100,150), 10)
    movieHole3.setFill("white")
    movieHole3.draw(win)
    
    movieHole4 = Circle(Point(150,100), 10)
    movieHole4.setFill("white")
    movieHole4.draw(win)
    
    movieHole5 = Circle(Point(185,185), 10)
    movieHole5.setFill("white")
    movieHole5.draw(win)
    
    movieHole6 = Circle(Point(115,185), 10)
    movieHole6.setFill("white")
    movieHole6.draw(win)
    
    movieHole7 = Circle(Point(185,115), 10)
    movieHole7.setFill("white")
    movieHole7.draw(win)
    
    movieHole8 = Circle(Point(115,115), 10)
    movieHole8.setFill("white")
    movieHole8.draw(win)
    
    # Makes sure the program runs until the user decides to quit
    while tryMovie == "yes":
        p = win.getMouse()
        while not clicked(clickCircle, p):
            p = win.getMouse()
        
        # Receives information from the text boxes
        inputBudget = float(inputBudgetBox.getText())
        inputCritic = float(inputCriticBox.getText())
        inputAudience = float(inputAudienceBox.getText())
        inputMovie = [inputBudget, inputCritic, inputAudience]
        
        sumOfBoxOffice = 0
        count = 0
        
        # Checks if the budget is less than 250 million dollars
        if inputBudget <= 250:
            
            # Finds the closest movies and gives it more weight than the next
            # closest
            similarMovies, nearestMovie = nearestMovies(inputMovie, movieDict, 20)
            for movie in similarMovies:
                sumOfBoxOffice += (2 ** (20 - count)) * movieDict[movie][3]
                count += 1
            estimatedBoxOffice = round(sumOfBoxOffice/2097151, 2)
        else:
            # Finds the three lines of best fit using the audience score, budget,
            # and critic score
            # When there are two capital letters, it is multiplying the 
            # corresponding values together, ie CB is critic score times budget
            similarMovies, nearestMovie = nearestMovies(inputMovie, movieDict, 1)
            sumOfBudget = 0
            sumOfBsquare = 0
            sumOfBB = 0
            sumOfCritic = 0
            sumOfCsquare = 0
            sumOfCB = 0
            sumOfAudience = 0
            sumOfAsquare = 0 
            sumOfAB = 0
            for movie in movieDict:
                sumOfBudget += movieDict[movie][0]
                sumOfCritic += movieDict[movie][1]
                sumOfAudience += movieDict[movie][2]
                sumOfBsquare += (movieDict[movie][0]) ** 2
                sumOfCsquare += (movieDict[movie][1]) ** 2
                sumOfAsquare += (movieDict[movie][2]) ** 2
                sumOfBB += movieDict[movie][0] * movieDict[movie][3]
                sumOfCB += movieDict[movie][1] * movieDict[movie][3]
                sumOfAB += movieDict[movie][2] * movieDict[movie][3]
                sumOfBoxOffice += movieDict[movie][3]
                count += 1
            aB = ((sumOfBoxOffice * sumOfBsquare) - (sumOfBudget * sumOfBB)) / ((count * sumOfBsquare) - (sumOfBudget ** 2))
            aC = ((sumOfBoxOffice * sumOfCsquare) - (sumOfCritic * sumOfCB)) / ((count * sumOfCsquare) - (sumOfCritic ** 2))
            aA = ((sumOfBoxOffice * sumOfAsquare) - (sumOfAudience * sumOfAB)) / ((count * sumOfAsquare) - (sumOfAudience ** 2))
            bB = ((count * sumOfBB) - (sumOfBudget * sumOfBoxOffice)) / ((count * sumOfBsquare) - (sumOfBudget ** 2))
            bC = ((count * sumOfCB) - (sumOfCritic * sumOfBoxOffice)) / ((count * sumOfCsquare) - (sumOfCritic ** 2))
            bA = ((count * sumOfAB) - (sumOfAudience * sumOfBoxOffice)) / ((count * sumOfAsquare) - (sumOfAudience ** 2))
            boxBudget = aB + (bB * inputBudget)
            boxCritic = aC + (bC * inputCritic)
            boxAudience = aA + (bA * inputAudience)
            
            # The budget is given more weight over the critic and audience scores
            estimatedBoxOffice = round((5/9 * boxBudget) + (2/9 * (boxCritic + boxAudience)), 2)
        
        # The calculated box office and related information is printed in the
        # graphics window
        results = "The estimated box office is \n$" + str(estimatedBoxOffice) + " million dollars"
        boxOfficeResults = Text(Point(800, 350), results)
        boxOfficeResults.draw(win)
        closestMovieStr = "The closest movie is " + str(nearestMovie) + "\nwith budget of $"+ str(movieDict[nearestMovie][0]) + " million \nand scores of " + str(movieDict[nearestMovie][1])+ " and "+ str(movieDict[nearestMovie][2]) 
        closestMovie = Text(Point(750, 250),closestMovieStr)
        closestMovie.draw(win)
        
        # Changes the color and funtion of the buttons to run again
        clickCircle.undraw()
        clickCircleText.undraw()
        clickCircle = Circle(Point(600,450),25)
        clickCircle.setFill("red")
        clickCircle.draw(win)
        clickCircleText = Text(Point(600,450),"Clear")
        clickCircleText.draw(win)
        
        # Allows the user to input new values
        p = win.getMouse()
        boxOfficeResults.undraw()
        clickCircle.undraw()
        clickCircleText.undraw()
        closestMovie.undraw()
        clickCircle = Circle(Point(600,450),25)
        clickCircle.setFill("green")
        clickCircle.draw(win)
        clickCircleText = Text(Point(600,450),"Run")
        clickCircleText.draw(win)
        
main()