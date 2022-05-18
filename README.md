# All Stock Analysis

# Overview of Project
## Purpose and Backgroun.
  ### Using VBA code to expand the dataset to include the entire stock market over the last few years. The purpose is to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

# Results
## The analysis is well described with screenshots and code.

### 1. Set tickerIndex equal to zero before looping over the rows.
#### Step 1a: Create a tickerIndex variable and set it equal to zero before iterating over all the rows. use this tickerIndex to access the correct index across different arrays.
![This is an image](https://github.com/Izzyycl/stocks-analysis/blob/main/Stock%20Analysis/a1.png?raw=true)

### 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
#### Step 1b: Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
##### The tickerVolumes array is Long data type.
##### The tickerStartingPrices and tickerEndingPrices arrays are Single data type.
![This is an image](https://github.com/Izzyycl/stocks-analysis/blob/main/Stock%20Analysis/1b.png?raw=true)

### 3. Access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
#### Step 2a: Create a for loop to initialize the tickerVolumes to zero.
![This is an image](https://github.com/Izzyycl/stocks-analysis/blob/main/Stock%20Analysis/2a.png?raw=true)
#### Step 2b: Create a for loop that will loop over all the rows in the spreadsheet.
![This is an image](https://github.com/Izzyycl/stocks-analysis/blob/main/Stock%20Analysis/2b.png?raw=true)

### 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
#### Step 3a: Find the total volume for the current ticker
#### Step 3b: Write an if-then statement to check if the current row is the first row with the selected tickerIndex
#### Step 3c: Write an if-then statement to check if the current row is the last row with the selected tickerIndex. 
#### Step 3d: Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
![This is an image](https://user-images.githubusercontent.com/87958408/168959787-b5d601ec-2634-4015-9256-cd2ff49903a0.png)

#### Step 4: Use a for loop to loop through arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
![This is an image](https://github.com/Izzyycl/stocks-analysis/blob/main/Stock%20Analysis/4.png?raw=true)

### 5. Code for formatting the cells.
![This is an image](https://github.com/Izzyycl/stocks-analysis/blob/main/Stock%20Analysis/5.png?raw=true)

### 6. There are comments to explain the purpose of the code.
#### Consistent Indentation, Avoid Deep Nesting, Control the line length, Comment the step on the code, Grouping.
### 7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.

#### Dataset Provided
![This is an image](https://github.com/Izzyycl/stocks-analysis/blob/main/resources/VBA_Challenge_2017.png?raw=true)
![This is an image](https://github.com/Izzyycl/stocks-analysis/blob/main/resources/VBA_Challenge_2018.png?raw=true)

#### VBA Solution
![2017](/Resources/2017.png)
![2018](/Resources/2018.png)

### 8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png.
![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

# Summary
## There is a detailed statement on the advantages and disadvantages of refactoring code in general.
### Advantages: 
  Easy to find error
  Make our code cleaner and more organizaed
  
### Disadvantages:
  Refactoring process can affect the testing outcomes.
  
## There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script.
### Advantages:
  Clearly organized, easy to find the code we need
  
### Disadvantages:
  Code might be duplicated
  Use more time and memory to run code
