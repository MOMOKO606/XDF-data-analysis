# Xi'an XDF data analysis program

This is a small project I made for the Foreign Examination Department of Xi'an New Oriental school.

The purpose is to enable teachers and hr to work more efficiently by data analysis. To be more specific, 

- We want to find the teachers who had too many and too few classes last year, then adjust the assignment of classes next year. 
  - We want to make sure the teachers with too many classes won't be too tired next year. 
  - Meanwhile, we want to find out the reasons that why some teachers had too few classes. Dig the reasons and help them improve, re-activate them.

- After adjusting the teacher’s schedule, calculate how many new teachers need to be recruited to meet the next year’s productivity target.



## Method

<u>**Input:**</u> 
" CLASS_SCHEDULE_FILE" from last year.

<u>**Output:**</u>
Report of teachers' performance 
Report of number of new teachers needed.

<u>**Process:**</u>
Define class as TpTrackers ( Teachers' performance Trackers ) to analyze " CLASS_SCHEDULE_FILE". Using time window and category to collect teachers' performance.
