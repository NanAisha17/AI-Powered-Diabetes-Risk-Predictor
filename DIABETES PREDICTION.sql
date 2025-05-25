CREATE DATABASE DIABETES_PREDICTION;

DROP TABLE DIABETES_DATASETS;

CREATE TABLE DIABETES_DATASETS (Gender VARCHAR(10),
								Age_Group VARCHAR(20),	
								Hypertension INT,	
								Heart_disease INT,	
								Smoking_history VARCHAR(30),	
								BMI FLOAT,	
								HbA1c_level FLOAT,	
								Blood_glucose_level FLOAT,	
								Diabetes INT);

BULK INSERT DIABETES_DATASETS
			FROM
				'C:\Users\owner\OneDrive\Desktop\NANA\darey.io\DIABETES PREDICTION DASHBOARD NEW.CSV'
				WITH (
						FORMAT = 'CSV',
						FIRSTROW = 2,
						FIELDTERMINATOR = ',',  
						ROWTERMINATOR = '\n',
   						TABLOCK);

SELECT * FROM DIABETES_DATASETS;

--HIGH RISK PATIENTS

SELECT Age_Group, 
		Gender, 
		BMI, 
		Blood_glucose_level,
		HbA1c_level
FROM DIABETES_DATASETS
WHERE BMI > 30 AND HbA1c_level > 7.0;

--DIABETES PREVALENCE BY AGE-GROUP

SELECT DISTINCT Age_Group,
				COUNT (Diabetes) AS DIABETES_COUNT
FROM DIABETES_DATASETS 
WHERE Diabetes = 1
GROUP BY Age_Group
ORDER BY DIABETES_COUNT
;

--SEGMENT BY RISK RANK

SELECT
		CASE
			WHEN BMI BETWEEN 18.5 AND 24.9 OR HbA1c_level < 5.7 THEN 'LOW RISK'
			WHEN BMI BETWEEN 25.0 AND 29.9 OR HbA1c_level BETWEEN 5.7 AND 6.4 THEN 'MODERATE RISK'
			WHEN BMI >= 30.0 OR HbA1c_level BETWEEN 6.3 AND 7.0 THEN 'HIGH RISK'
			ELSE 'VERY HIGH RISK'
		END AS DIABETES_RISK,
	Gender,
	Age_Group,
	Hypertension,
	Smoking_history,
	BMI,
	Blood_glucose_level,
	HbA1c_level
FROM DIABETES_DATASETS
GROUP BY Gender,
	Age_Group,
	Hypertension,
	Smoking_history,
	BMI,
	Blood_glucose_level,
	HbA1c_level
ORDER BY DIABETES_RISK;