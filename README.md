# Excel Projects

This repository contains my Excel references and projects. 

### File descriptions

#### Excel_formulas_charts_reference
This Excel file contains detailed explanations and examples of the formulas, functions, charts, and graphs I have learned in [Maven Analytics](https://www.mavenanalytics.io) Excel courses. 

#### Austin_Animal_Center_Intakes_Dashboard_v2
This Excel file contains the raw data, analysis and transformation steps, and an interactive dashboard, created with Austin Animal Center's intake data from October 2013 through October 2022. To use the interactive drop-down menus, download the Excel file. Note that I did remove fields that were not used in dashboard from the raw data tab in order to make the file small enough to upload to this repository. To view the entire dataset, which is updated on a daily basis, visit the [City of Austin Open Data Portal.](https://data.austintexas.gov/Health-and-Community-Services/Austin-Animal-Center-Intakes/wter-evkm)

#### COA_Annual_Crime_Data_2015-2018
This Excel file contains two interactive graphs showing crime trends for the City of Austin from 2015 through 2018. The first sheet, "Trends by crime type," contains a line graph showing trends by month and type of crime for each year, 2015-2018. The second sheet, "Location," is a heat map showing the distribution of crime by zip code. The darkest green color represents the lowest crime numbers, and the darkest red represents the highest crime numbers. Download the files to use the interactive drop-down menus on both sheets. 

The remaining sheets in this file contain the raw data used in the analysis and to create the graphs. Note, I did remove fields that I did not use to create the graphs from the data. The datasets were downloaded from the City of Austin Open Data Portal on November 20, 2022. See below for a link to view each original dataset. 

[City of Austin Crime Data: 2015](https://data.austintexas.gov/Public-Safety/Annual-Crime-Dataset-2015/spbg-9v94)

[City of Austin Crime Data: 2016](https://data.austintexas.gov/Public-Safety/2016-Annual-Crime-Data/8iue-zpf6)

[City of Austin Crime Data: 2017](https://data.austintexas.gov/Public-Safety/2017-Annual-Crime/3t4q-mqs5)

[City of Austin Crime Data: 2018](https://data.austintexas.gov/Public-Safety/2018-Annual-Crime/pgvh-cpyq)

#### pizza_order_dashboard
This Excel file is an interactive dashboard that allows the user to see the grand total for a "pizza order" by selecting a pizza category, type of pizza, size, and quantity. The dataset, "Pizza Place Sales," contains data from a fictitious pizza place and was obtained from [Maven Analytics' Data Playground.](https://www.mavenanalytics.io/data-playground) I used several advanced functions and formulas to create this dashboard, including VLOOKUP, XLOOKUP, SUMPRODUCT, FILTER, conditional formulas, and data validation techniques to create dynamic drop-down menus. As with all of my other Excel files, download the file to use the interactive features. 

#### unicorn_pivot
This Excel file contains several pivot tables I created to explore an interesting dataset about private companies with a valuation of over $1 billion as of March 2022, a.k.a, "Unicorn Companies." This dataset is also retrieved from [Maven Analytics' Data Playground.](https://app.mavenanalytics.io/datasets?page=2) 
In the raw data (last sheet in the file), the new columns I created to reformat some of the fields have a light yellow background on the header. 
Here is a brief summary of what each pivot table shows and some interesting insights I gained as a result of using pivot tables to analyze this data (as always, download the file to explore for yourself!)
* ROI: This pivot table and pivot chart show the return on investment for the companies, based on the funding when they were founded and the total valuation as of 2022. I applied a value filter for top ten to make the pivot chart more readable. The table and chart can be filtered by country and industry. The filter is currently set to exclude the United States because one company, Zapier, a U.S. internet and software company, has such a high ROI that it completely dwarfs every other company on the pivot chart!
* Years to Unicorn: The first pivot table on this sheet shows the average number of years it took a company to reach unicorn status, broken down by industry and year. Although the same information is in the first pivot table, I made a second table to show the average number of years to reach unicorn status overall for each industry. The first pivot table shows an interesting trend that is pretty consistent across industries: Companies founded more recently reach unicorn status more quickly, on average, than companies founded earlier. The second pivot table shows that companies in the healthcare industry take the longest average number of years to reach unicorn status. It should be noted that this industry has a major outlier: Otto Bock Health Care, founded in Germany in 1919, with $0 funding, took 98 years to reach unicorn status and has a valuation of $4B as of 2022!
* Hubs: This table shows the number of companies and percentage of the industry by city, sorted in descending order by percentage of the industry. Beijing, San Francisco, and New York appear to be hubs for several industries.
* Top performers: The table shows the top ten countries with the most unicorns, and the chart shows the top 10 highest valued companies overall. As you can see, the United States has more than three times as many unicorns (562) as the country with the second most, China (173); however, China has the top two highest valued companies while the United States takes 3rd and 4th place. 
* Top investors: I eventually plan to revisit this table and do more with it, but for now, it shows that very few companies have a 4th investor. To create this view, I first used the TEXTSPLIT function to create new fields in the raw data, one for each investor. 
