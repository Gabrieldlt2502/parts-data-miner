# Automotive Parts Compatibility Miner

## Business Context
In the auto parts industry, a single SKU can fit hundreds of different vehicles. Manually mapping these "compatibilities" is nearly impossible at scale.

##  The Solution
I built a robust data mining tool that automates the extraction of application data from industry-standard databases. 
* **Data Transformation:** Converts unstructured web tables into clean, relational Excel data.
* **Anti-Throttling logic:** Implemented adaptive sleep cycles and human-like browsing patterns to maintain connection stability.
* **Incremental Progress:** Includes an "Autosave" feature that commits data every 10 SKUs to prevent loss during large-scale mining operations.

##  Tech Stack
* **Python / Selenium:** For complex web navigation and dynamic content rendering.
* **Pandas:** For data structuring and cleaning.
