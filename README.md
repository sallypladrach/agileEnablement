## What is this? 
This script is something that you can use to create an automation with Google app scripts that will create a customized slide deck when the user fills out a simple form. 

## How do I use it? 
More detailed instructions will be added to the Read Me later, but for now, here's how you can get started. 

1. Create your template deck - This should have all of the slides that could possibly included, and placeholders where you may want custom text to appear. Placeholders in the slide deck should be inside of these
   brackets, for example {{client name}}

2. Give your template deck tags in the notes - each thing that you want to use to dictate with slides should be included from the master template should have tags that will appear in the notes section. For example, if I 
want to customize by industry, the slides that are industry-specifc should have the names of industries in the notes section, like "logistics" or "education". Any slides that you want to have in every deck (like an intro
or contect slide) should have "all" in the notes.

3. Create your Google form, ensuring that the choices you've input will exactly match the words useds as tags on your slides. For example, if I put "education" on the education-specific slides I want to include, but my form choice says "K-12 Ed", the script will not know to include the slide because the field on the form does not match the tags in the slides.

4. Once you've set up the form, go to the responses tab in the form settings and create the sheet that responses will flow to when filled out. 

5. Go to the sheet and go to Extensions > Apps Script in the menu

6. Copy-paste the code from this repo into Google App scripts and replace the appropriate variables for your setup, including the name of the sheet, the column indexes, where the presentations should go, etc.

7. The names of the variables are from my own project, but you can adjust them to fit your form if you want, just be sure to change the variable names throughout your code to match.

8. Remember to set up your triggers in Google app scripts to run the script each time a new row is added to the sheet.  

9. Ensure there's at least one row of data in your form, and run the script to test it out. If you need help or would like to help me make this more user-friendly by letting me watch you set it up, reach out to me here or on [LinkedIn](https://www.linkedin.com/in/sallyladrach/)
