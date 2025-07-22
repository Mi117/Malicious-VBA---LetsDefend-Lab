# Malicious-VBA---LetsDefend-Lab
Malicious VBA — LetsDefend Challenge WriteUp

## Overview

In this lab we will perform a deep analysis on a suspicious document attachment in an email. The document has been flagged suspicious as it presents some strings from the VBA Macro document and they are obfuscated.

Link to the Lab: https://app.letsdefend.io/challenge/Malicious-VBA

-------------------

## What is Virtual Basic Applications (VBA)?

Visual Basic for Applications (VBA) is a scripting language built into Microsoft Office applications, enabling users to automate tasks through macros. While VBA is a legitimate tool for productivity, it is frequently abused by threat actors to deliver and execute malicious payloads. Malicious VBA macros are often embedded in Office documents (e.g., .docm, .xlsm) and triggered automatically when opened, exploiting social engineering to bypass user awareness. These macros can download malware, execute system commands, or establish persistence, making VBA a common vector in phishing and initial access attacks. Its deep integration with Office and the Windows environment makes VBA a potent threat in modern cyber campaigns.

## How attackers can exploit VBA?

Attackers exploit VBA by embedding malicious macros into Office documents — typically Word or Excel files — and distributing them via phishing emails or malicious downloads. These documents often use social engineering tricks (e.g., fake invoices or job offers) to convince the victim to enable macros, which are disabled by default in modern Office versions for security.

#### MACROS: small programs or scripts written in Visual Basic for Applications (VBA) that automate tasks within Microsoft Office applications like Word, Excel, PowerPoint, and Access. They allow users to perform repetitive actions — like formatting data, generating reports, or updating documents — with a single click or automatically when a file is opened.

https://support.microsoft.com/en-us/office/introduction-to-macros-a39c2a26-e745-4957-8d06-89e0b435aac3

Once enabled, the VBA macro executes automatically using event triggers like AutoOpen() or Workbook_Open(). The code can then:

- Download and execute malware from remote servers
- Create files or registry keys for persistence
- Run PowerShell or CMD commands to control the system
- Exfiltrate data or credentials
- Disable security tools or modify system settings
- Because VBA runs within a trusted Office application, it can bypass some security controls, making it a favored tool for initial access in many malware campaigns, including ransomware, info-stealers (like Agent Tesla), and remote access trojans (RATs). Obfuscation techniques like hex encoding or dynamically created objects are often used to evade detection by antivirus software.

### The subtle art of Obfuscation
Obfuscation is a technique commonly used by attackers to hide the true purpose of their malicious code and make it harder for analysts or security software to detect or understand it. In the context of VBA exploit, this often means disguising important strings — like URLs, file names, or commands — by encoding them in hexadecimal or splitting them into meaningless variables. Instead of writing the actual malicious command directly in the code, the attacker breaks it into pieces or hides it behind functions that decode it only when the macro runs. This makes the code look less suspicious at first glance and helps it slip past antivirus programs.

#### Essentially, obfuscation is like writing in a secret code — it doesn’t change what the malware does, but it makes it much harder to read and recognize.

--------------------------

## Scenario of the Challenge

One of the employees has received a suspicious document attached in the invoice email. They sent you the file to investigate. You managed to extract some strings from the VBA Macro document. Can you refer to CyberChef and decode the suspicious strings?

Please, open the document in Notepad++ for security reasons unless you are running the file in an isolated sandbox.

Malicious Macro: /root/Desktop/ChallengeFiles/invoice.vb

This challenge prepared by @RussianPanda

-----------------

## Walkthrough

### Q1) The document initiates the download of a payload after the execution, can you tell what website is hosting it?

We start our investigation by accessing the file’s folder — /root/Desktop/ChallengeFiles/invoice.vb — and perform a primary analysis, extracting the file hash to then cross-reference on popular malware databases (like VirusTotal in this case).

<img width="914" height="699" alt="1-file" src="https://github.com/user-attachments/assets/7d5ace04-b785-4a50-ade2-fd6382b4ccce" />
<img width="911" height="304" alt="2-file-hash" src="https://github.com/user-attachments/assets/1a3e0ae1-f55d-4e4d-a7c7-a01702bb1a09" />
<img width="1923" height="835" alt="3-malicious nature virus total check" src="https://github.com/user-attachments/assets/4ab99a09-f767-42eb-81cc-13477e73586b" />

From the analysis the file is flagged as part of phishing processes. Once run, the file invoice.vb attempts to contact a web server for downloading the malicious payload — in this case https://tinyurl.com/g2z2gh6f

### Q2) What is the filename of the payload (include the extension)?

We then proceed to analyze the file, using our trusted Notepad or TextEditor and we get this:

<img width="1918" height="1080" alt="4-opening file with notepad" src="https://github.com/user-attachments/assets/c1b2bcf1-6b35-4495-8b7a-1e274c17a611" />

We can clearly see that many variables are split into hexadecimal code, which at first glance may seems meaningless, but it might hide malicious commands. To decode it we leverage the powerful tool of Cyberchef — a free, open-source web application that offers a wide range of operations for encoding, decoding, encrypting, decrypting, and analyzing data.

By starting to decode the hexadecimal values of the strings, we start to reveal valuable information, as well as our answer:

<img width="1915" height="1075" alt="answer 1 - method 2" src="https://github.com/user-attachments/assets/46f66101-c341-4aa5-b25d-43dadd4daa6f" />

_(Answering previous question 1 - decoding hexadecimal values)_

<img width="1917" height="1079" alt="answer 2" src="https://github.com/user-attachments/assets/ea67acac-59c3-4ff5-880c-e44248296b42" />

_(Answer of Question 2) _

The payload is **dropped.exe**

### Q3) What method is it using to establish an HTTP connection between files on the malicious web server?

We keep on decoding the values to find the next answer:

<img width="1889" height="1029" alt="answer 3" src="https://github.com/user-attachments/assets/a6cee8e1-942c-48d6-a1d4-b8fd05c04384" />

This object is a built-in Windows component that allows VBA scripts to send HTTP or HTTPS requests to remote servers — just like a web browser, acting as a silent and hidden downloader in the background and without user interaction.

### Q4) What user-agent string is it using?

<img width="1892" height="1077" alt="answer 4" src="https://github.com/user-attachments/assets/117d19d8-544c-4631-8172-285b524df3c1" />


### Q5) What object does the attacker use to be able to read or write text and binary files?

<img width="1903" height="1068" alt="answer 5" src="https://github.com/user-attachments/assets/9c178a4d-7856-4d16-8b8c-37c1583678af" />

**ADODB.Stream **is a built-in Windows object that allows a script (like VBA) to read, write, and save data to and from a stream, such as a file or binary content downloaded from the internet. A stream allows programs to read or write data in chunks, making it a more efficient process, especially for large files or binary content.

### Q6) What is the object the attacker uses for WMI execution? Possibly they are using this to hide the suspicious application running in the background.

<img width="1912" height="1062" alt="answwr 6" src="https://github.com/user-attachments/assets/cf0c1e71-f145-4c8b-9a6e-739b00325c9e" />

This line creates a connection to the WINDOWS MANAGEMENT INSTRUMENTATION (WMI) — a built-in feature in Windows that lets programs and scripts interact with the operating system to manage and control various parts of the computer, like running processes, checking system info, or changing settings — service and specifically accesses the Win32_Process class, which allows the script to create and run processes on the system.

_Check the in-depth WriteUp on Medium:_https://medium.com/@AtlasCyberSec/malicious-vba-letsdefend-challenge-64b3676d887f

-----------------

## Closing Thoughts
From this lab, I learned how malicious VBA macros embedded in Office documents can be used to silently download and execute malware on a victim’s system. The macro leverages obfuscation techniques — like hex-encoded strings and confusing variable names — to hide critical details such as URLs and file paths, helping it evade detection. It also uses Windows-native components, specifically MSXML2.ServerXMLHTTP.6.0 to establish HTTP connections and download the payload, and ADODB.Stream to write the malicious file to disk. For execution, the macro employs WMI (Win32_Process) to launch the payload stealthily with no visible window, falling back on WScript if needed.

This lab demonstrated how attackers use built-in Windows features and obfuscation together to create effective and hard-to-detect malware delivery mechanisms within seemingly innocent documents.
As I continue my journey into cybersecurity and blue team operations, labs like this one deepen my analytical mindset and provide valuable, real-world experience in dissecting malware artifacts — preparing me for future roles in threat detection, incident response, or SOC analysis.

I hope you have found this Write-Up insightful and make sure to follow me on all the platform to stay up-to-date with my latest analysis, write-ups and be part of my journey into the world of Cybersecurity!

### Link: https://linktr.ee/atlas.protect
