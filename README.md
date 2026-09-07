# nassp-checklist-annotator
This idea came about as I was finding myself often pausing NASSP to go and look up Verb/Noun/Program/DEDA lists and what they mean and what units they use.
My boundaries for this project were, only use official definitions the pilots had with them, do not replace any original entries, only append annotations.

I originally created this solution just for me, and I now think other NASSP users might benefit from this work.
At first, it was just a simple manual find/replace action in excel which got very repetitive.
I was now in search of a way to automate the find/replace action in excel with a view of one day applying these annotations to any NASSP mission that has XLS checklists.

Here is the approach I settled on for CSM:
Build a dictionary file with all of the CSM Verb/Noun/Program codes and descriptions that Mike Collins had with him in the CSM on Apollo 11.

CSM Source:
Apollo 11 CSM Operations Checklist SKB32I00080-307 found at:
https://www.apollojournals.org/afj/ap11fj/a11-csmocindex.html

Here is the approach I settled on for LM:
Build a dictionary file with all of the LM Verb/Noun/Program/DEDA codes and descriptions that Neil and Buzz had with them in the LM on Apollo 11.

LM Source:
APOLLO XI LM-5 FLIGHT CREW G&N DICTIONARY
https://www.ibiblio.org/apollo/Documents/ApolloXI_FlightCrewG&NDictionary_5-29-1969.pdf

Note: For the DEDA inputs/outputs, there were some items not shown in the G&N Dictionary (less than 10), and therefore the following LM/AGS source was used to fill in any gaps:

LM/AGS OPERATING MANUAL FLIGHT PROGRAM 6
https://apollojournals.org/alsj/LM_AGS_FP6_OperatingManual.pdf

Once the respective CSM and LM dictionaries were built, I intended to be the least intrusive I could be to the source checklist files in the older XLS format.
The method I settled on was to open both my dictionary file that I created in excel and the NASSP checklist file in excel and use several VBA scripts to annotate extra information to the end of the original text in the NASSP checklist cells when they match verb/noun/program/DEDA codes found in the dictionary.

The advantages of this approach is that the source XLS file is not excessively modified, and that the original text and annotations are retained as the original checklist dev intended.
Also, I really do not want to be responsible for missing important existing annotations like these examples:

CSM Checklist:

V75 (No Enter)

F 06 89 (LAT, LONG/2, ALT) (Do Not PRO)

After my VBA scripts run, this is what those same cells look like:

V75 (No Enter) - Backup Liftoff

F 06 89 (LAT, LONG/2, ALT) (Do Not PRO) - Display Decimal - Landmark, R1=Lat (+ North), R2=Long/2 (+ East), R3=Alt (.001° .001° .01 NM)

LM Checklist:

223+00020 (Do Not Enter)

310+02600 (No Enter)

After my VBA scripts run, this is what those same cells look like

223+00020 (Do Not Enter) - Altitude Update Input (100 ft)

310+02600 (No Enter) - Targeted TFI TPI (For TPI Search Routine) (.01 min)

As you can see, the only major disadvantage is that the text is longer, however personally, I would rather have more information than none, especially having units in parentheses at the end makes a big difference for me.
In future, depending on feedback, I could potentially improve this long text by changing the VBA to only apply annotations when no original annotation exists in the checklist.
For now, I am using the extra information as both a sanity check and to verify if the original annotation is correct.

The vba scripts also create logs on your desktop of exactly what was modified in the following format.

Tab | Original text -> Original text - Annotation1 - Annotation2 (Annotation3)

Example log lines:

Flightplan | V37E 52E -> V37E 52E - Change Program - IMU Realign

Flightplan | V77E -> V77E - Set Rate Command/Attitude Hold Mode in DAP

Flightplan |   F 37 ->   F 37 - Change Program

Rendezvous |   F 06 47 ->   F 06 47 - Display Decimal - R1=LM Weight, R2=CSM Weight (lbs)

Flightplan | 400+0 -> 400+0 - Attitude Hold

DOI | 400+1 -> 400+1 - Auto Guidance Steering

Flightplan | 400+2 -> 400+2 - Z-Body Axis Steering

DOI | 400+3 -> 400+3 - IMU Align

One thing to be aware of is that where the source material uses the delta Δ character, the logs show a ?, however the XLS checklist is written with the correct delta character.

Anyway, this is not a serious thing, just a "problem in search of a solution" exercise I gave to myself that I thought I would share and try to contribute to the NASSP community in whatever small way I can.
The goal was to try and retain the realism that NASSP has while also helping to educate the user about what inputs/outputs they are performing as they progress through the checklists.
