<h1> <img src="https://raw.githubusercontent.com/JoseChirif/Mass-file-renaming-with-excel/refs/heads/main/assets/icon.png" width="20" height="20" />  MASS FILE RENAMING WITH EXCEL </h1>


Generate an automatic excel template with the name of all files in the same folder of your project. Then, rename all of them with one click (as the new value you assing in the excel).

![Mass file renaming intro](https://github.com/user-attachments/assets/a83851ec-b4ee-43c1-a433-60317cde5f2f)

  <!--- Badges /> --->
<div style="display: flex; justify-content: center; gap: 20px; align-items: center;">
  <img src="https://img.shields.io/github/languages/top/JOSECHIRIF/Mass-file-renaming" alt="Languages" />
  <img src="https://img.shields.io/badge/python-3.11.5-blue" alt="Python version" />
  <img src="https://img.shields.io/github/license/JoseChirif/Mass-file-renaming-with-excel" alt="License" />
  <img src="https://img.shields.io/github/release/JoseChirif/Mass-file-renaming-with-excel" alt="release" />
</div>

<br>


# 🛠️ SETUP
**Steps to setup the program:**

1. Download the last release.rar

<table>
  <tr>
  <td><img src="https://github.com/user-attachments/assets/b34da49e-b074-4c95-a122-f77645adba06" alt="go to release section" /></td>
  <td><img src="https://github.com/user-attachments/assets/6f47e0fd-77bd-46be-b62e-c9e666084d09" alt="download the last rar file" /></td>
  </tr>
</table>

2. Extract the .rar program

3. Install virtual environment with command `python -m venv venv`

4. Activate the virtual environment with `venv\Scripts\activate`

5. Install dependencies from requirements.txt: `pip install -r requirements.txt` (check that your venv is active).

<br>

**Then you have 2 options:**
  
<!-- Option a: Make a .exe file - RECOMMENDED (Click here):START -->
<details>
  <summary>
    Option a: Make a .exe file - RECOMMENDED (Click here)
  </summary>

  <ol start="6">
    <li>Run: <code>pyinstaller --onefile --windowed --clean --noupx --add-data "assets/*;assets" --add-data "config/*;config" --add-data "functions/*;functions" --add-data "locales/*;locales" --add-data "src/*;src" --add-data "LICENSE.txt;." --add-data "README.md;." --icon "assets/icon.ico" --name "0 rename.exe" run.py</code></li><br>
    <li>Then the folder dist will be created in the project's folder. Inside is the exe file.
  </ol>

  <img src="https://github.com/user-attachments/assets/d6bbdbcc-7d98-4fc1-a4cc-5d4c4fff7511" alt="dist folder">

</details>
<!-- Option a: Make a .exe file - RECOMMENDED (Click here):END -->




<!-- Option b: Execute from python (Click here):START -->
<details>
  <summary>Option b: Execute from python (Click here)</summary>

  <ol start="6">
    <li>Execute run.py to enter the main menu: <code>python run.py</code></li><br>
  </ol>
  
  <img src="https://github.com/user-attachments/assets/5550e35f-2d79-4afd-bb6a-b61d27045e82" alt="father's directory">

  <!---
  <img src="https://github.com/user-attachments/assets/efa73448-b450-4e51-a26c-c2f8ebb882a7" alt="father's directory">
  --->

  <p><strong>Scripts will modify the name of files in project's parent directory instead of where the executable is.</strong></p>

</details>
<!-- Option b: Execute from python (Click here):END -->    


# 🌐 LANGUAGES (instructions):
- [Instructions in **ENGLISH** (click here)](#-instructions)

- [Instrucciones en **ESPAÑOL** (click here)](https://github.com/JoseChirif/Mass-file-renaming/blob/main/instructions/Instructions%20-%20es.md)


# 📑 INSTRUCTIONS
  1. Move the executable to the folder where the files you want to modify are located.
    ![move_to_folder](https://github.com/user-attachments/assets/a186ba66-b2f7-452f-8797-4f054907d76f)

  2. With the main menu open you can run all scripts.
    ![main menu](https://github.com/user-attachments/assets/9e9d432d-d3e2-4f8a-a4c4-dea361f054b3)

  3. The first button will create an excel with the name of all files and folders of the folder where it is.
    ![excel template](https://github.com/user-attachments/assets/c43eb533-498d-46a5-87d3-1ab98e0f8348)

  4. Fill column C of the excel with new names (or leave it empty if you don't wanna change the name).
    ![Flash fill excel](https://github.com/user-attachments/assets/ec5e8c1a-dc87-49f7-bff6-abe98b32a57c)
    **Tips:** 
    - If there is a regex, made 2 or 3 examples and use excel flash fill.
    - If you want to unlock the sheet to use more columns, click on the 3rd button of the main menu does it.

  5. The second button will change the original name of all files with the new name assings in the column C of the excel.
    ![renaming](https://github.com/user-attachments/assets/e8aa9663-363b-4297-aa6f-55cae6d83c77)


  **You have 2 options:** <br>
    ![options to modify names](https://github.com/user-attachments/assets/8d4136fe-5dc2-43c5-875a-fc729e16124d) 
  - Option 1: Made a new folder (with the same name as the project) and copy all the files with the new name.
  - Option 2: Overwrite the files with the new names.


**DONE, you made it well!**



# ✍️ Author
[@Jose Chirif](https://github.com/JoseChirif)

## 🚀 About me
I'm an Industrial Engineer specialized in process optimization, business intelligence and data science.
[Porfolio - Network - Contact](https://linktr.ee/jchirif)








