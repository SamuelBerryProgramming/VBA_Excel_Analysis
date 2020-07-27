# Analysis on Stock and Script Performance
## Purpose
  A macro was written to perform analysis of total volume and yearly return on two different years of data for twelve different stocks. The original code used a nested for loop to iterate through all the rows twelve times, once for each of the different stocks. The code was then refactored to loop through all the rows a single time, keeping track of each stocks ticker. After refactoring, modest performance gains were obtained. 
  
## Results
#### Stock Performance
  After running the macro on each year of data, it can be seen that in 2017 all of the stocks except TERP saw at least some increase in value. In 2018 however, all stocks except ENPH and RUN saw declines in value.
#### [Stock Performance 2017]()
#### [Stock Performance 2018]()

#### Code Performance
  At runtime, both the original and refactored code acheived the same numerical results from the data set. The original code utilized two for loops, one nested into the other. This resulted in iteration over all rows twelve times, once for every stock ticker. The original code took slightly longer than one second when ran on each year, while the refactored code took approximately 0.17 seconds to complete. The refactored code uses a single for loop to iterate through the rows once, while using three arrays and an index to keep track of each stock ticker. The new code ran nearly six times faster than the original.
#### [Original Code Performance 2017]()
#### [Original Code Performance 2018]()
#### [Refactored Code Performance 2017]()
#### [Refactored Code Performance 2018]()

## Summary
#### Advantages and Disadvantages of Refactoring Code
  Some advantages of refactoring code are possible performance improvements, better maintainability, and increased reusability. This is acheived by going through old code and compartmentalizing the logic into units, making descriptive comments, and reducing the amount of code used to do the work. Possible disadvantages are the amount to time consumed dissecting the original code, reassembling the code, and debugging it all. After all that, it is still possible for the code to run in a new, unexpected way that introduces defects in the new solution. Returns are also diminished based on how well the code ran in the first place.
  
#### Original vs. Refactored VBA
  Considering these positive and negative aspects of refactoring, the largest benefit this VBA script received was an increase in performance. While the times to completion were perceptually negligible when ran on this data set, the refactored code would be much quicker when ran on larger data sets than the original. Even though maintainability was not increased, it became much more viable to use this script in other circumstances. The drawback was the amount of time it took to change this code which was already working just fine for its application. 
