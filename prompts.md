# A feladat

Bluetoothos hőmérők letöltött adatainak összegyűjtése és vizualizálása, kimutatások készítése. Lásd specification.md.

# Tanulságok

- Use VS Code Tasks - Press Ctrl+Shift+P → "Tasks: Run Task" → Choose from:
- C# példa kód a timestamp kezelésre segített neki.
- Egy idő után kezd belassulni... és mintha ő is próbálná kezelni ("summarizing conversation history...").
- Ha az egyes **feature-öknek van kódja**, akkor azt tényleg használja is a hivatkozásokhoz. Ezeket a specifikációban így vezettem be a fejezet elején: "In this section, I will assign an identifier for every function so that we can easier talk about them. The identifier will be before the title of the function, in brackets and may contain letters and numbers. For example "(IN001)", or "(STAT003)"."
- Hasznos lehet bevezetni a **"Q:" prefix** mellett (csak kérdés, ne módosítson semmit) egyfajta **munkafázis kéréseket**, hogy pl. ellenőrzötten csak akkor töltsön időt doksi frissítéssel, amikor kell.
    - DOC: update @spacification.md and @InformationForAI.md if needed.
    - TEST: run unit tests. Fix issues and update unit tests if needed.
    - REFACTOR

# Konkrét promptok

I want to create an appliation which reads CVS files packed in zip files, loads data from them and generates nice diagrams and maybe excel tables. I want to create a nice set of unit tests for this as well. Is a python console application a suitable environment for that?

(The agent started to implement and recognized the spacification.md I was working on in parallel. It created test input based on the format description and unit tests, and run them.)

---

Now I have copied the true input data zip files into the data folder with names TempLogs*.zip . Please have the application process them and write the results into its data storage JSON file. Have a look at @specification.md to see what I mean.

---

I have a C# method to convert those timestamps to DateTime objects. Have a look at it and implement the import accordingly:

       private DateTime GetDateTimeFromCsvTimestamp(long csvSecondsTimestamp)
        {
            // Convert seconds to ticks
            long timestamp = csvSecondsTimestamp * TimeSpan.TicksPerSecond;

            // Unix epoch in ticks
            long unixEpochTicks = 621355968000000000;

            // Create DateTime instance
            var dateTime = new DateTime(unixEpochTicks + timestamp, DateTimeKind.Utc);

            return dateTime;
        }


(Ebből rájött, hogy nem Excel timestamp van, hanem Unixos!)
---

First I have updated @specification.md as some devices may record with 30 minutes interval, so accept continuous recording if the time difference between the entries is less than 35 minutes.

----

We are in a git repository. Please create a .gitignore file.

---

Now lets resume the development of the nice visualizations you proposed.

(Itt sikerült valahogy \n problémákba szaladni és utána újra kellett kezdeni a vizualizációs fájl generálását...)
---
You seem to be stuck again. Retry the debugging of the visualization.

---
It seems we have lots of problems with the bad timestamps. Extend the @specification.md for importing new zip files to skip records with timestamp before 2020 at all.

(Az impementációt is elkezdte magától...)

## Nagy újrakezdés a visualization témában

I rolled back some changes and I start a new chat as you got confused and slow. Have a look at the @specification and start to implement simple visualization features.

(Erre újra végignézte a kódot, egész ügyesen. És megint gyorsnak látszik.)

------

I have removed the database JSON as it was outdated.
Please check the @data_importer.py, according to @specification.md it should not import data records with timestamp before 2020. Is that implemented? If not, do it. And then rerun the script.

(Erre most már jól működik, magától le is tesztelte.)

----

Please update @specification.md to make it better and add everything missing from it based on the implementation.

(A feature kód konvenciót követve további terveket is javasolt!)

## Jegyzeteljen magának

If we start a new chat, you have to read the codebase again. To make understanding the codebase faster, create a file named "InformationForAI.md" where you can make any kind of notes for yourself. The goal is to make your catch-up faster if this file is there.

## Új adatok letöltése után frissítés

Now I have downloaded new measurement data so please update the database based on the ZIP files and rerun all the statistics.

## Question mode

From now on if I tart a prompt with "Q:", only answer my question but do not take any further action like changing source code.

## Guessing ventillation status: BE, Kek and Terasz

Q: There is an assumption in @specification.md at feature STAT002. What kind of statistic could support this assumption?

---

Q: Unfortunately we do not know when the ventillation is ON or OFF. What kind of statistic could guess its status?

(Ilyeneket javasol, mint "Heat Transfer Response Detection"!)

-----

Help me implement suggestion 5, temperature gradient analysis to estimate ventillation status (on or off).

Ilyenekre jön rá: "The analysis shows that we don't have enough synchronized data points. This could be because the devices don't record at exactly the same times. Let me modify the synchronization logic to be more flexible by allowing some time tolerance:"

"Let me also add some debug information to see how many synchronized points we're getting:"

---

It is hard to see when we have confident estimates. Create also an image where rows represent date (each row is a day) and the columns represent the hours. The color of each cell (pixel) corresponds to the ON/OFF/Unceratin status.

(Válaszban: "Excellent! Now let me also update the InformationForAI.md file to include information about these new heatmap visualizations:")

---

Create an additional heatmap simply showing the temperature difference between the room (T3_Kek) and the intake T1_BE. Use a termal colormap to visualize the temperature difference.

----

Based on my observations, @ventilation_temperature_difference_heatmap.png is the only useful reasult from STAT002. Please remove the other results to simplify the codebase.

## GUI készítés

Now lets create a separate executable console application which will have a GUI: it shows diagrams of temperature values of selected sensors. Sensors can be selected with checkboxes and the starting and ending date of the diagram can be chosen with date pickers. Prepare the code to visualize time series data generated by statistics, like the temperature difference in STAT002.

---

Add a derived statistic to show the temperature difference "T3_Kek - T2_Terasz". Implement it in the GUI so there is no need to run anything before using it.

----

Modify STAT002 visualization in the GUI so that it does not require running anything before the launch of the GUI.
