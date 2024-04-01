import pandas as pd 
import openpyxl
from openai import OpenAI
import openpyxl

client = OpenAI(api_key=GPT_API_KEY)

# Create a new Excel workbook
workbook = openpyxl.Workbook()

# Select the active sheet (by default, it's the first sheet created)
sheet = workbook.active

response = client.chat.completions.create(
  model="gpt-3.5-turbo",
  messages=[
    {
      "role": "system",
      "content": "You will be provided with unstructured data, and your task is to parse it into CSV format. each cell should contain 1 question regarding the process \n\nQuestions must be in format of: What does the process concern about ? Conditions of the process : Steps of the process : \n\nthe word process must contain the name of the process mentioned in the text , for every question there must be an answer , all the question - answer must be in greek , format question - answer :  question answer "
    },
    {
      "role": "user",
      "content": " Βεβαίωση αποδοχών τακτικού / εποχικού προσωπικού ΕΛΓΑ Η διαδικασία αφορά στην χορήγηση βεβαίωσης αποδοχών τακτικού / εποχικού προσωπικού ΕΛΓΑ. Όσοι ανήκουν στο τακτικό και εποχικό προσωπικό του Οργανισμού Ελληνικών Γεωργικών Ασφαλίσεων (ΕΛΓΑ) μπορούν να εκδώσουν βεβαίωση αποδοχών. Η βεβαίωση αφορά συγκεκριμένο οικονομικό έτος (από το 2011 και μετά) και εκδίδεται για φορολογική χρήση.Προϋποθέσεις:1.Κατοχής κωδικών για είσοδο σε λογισμικό:Να διαθέτει ο χρήστης κωδικούς Taxisnet.1.Επαγγελματικές:Να ανήκει ο χρήστης στο τακτικό προσωπικό του ΕΛΓΑ.1.Επαγγελματικές:Να ανήκει ο χρήστης στο εποχικό προσωπικό του ΕΛΓΑ.Ψηφιακά βήματα:1.Αυθεντικοποίηση χρήστη με κωδικούς Taxisnet:ο χρήστης εισάγει τους προσωπικούς κωδικούς που διαθέτει στο Taxisnet1.Αποτυχία αυθεντικοποίησης χρήστη:Σε περίπτωση λανθασμένης καταχώρισης του ζεύγους των κωδικών, αποκλείεται από το σύστημα, η είσοδος στην υπηρεσία.1.Επιτυχής αυθεντικοποίηση χρήστη:Εισάγεται επιτυχώς το ζεύγος κωδικών Taxisnet και ο ενδιαφερόμενος αποκτά πρόσβαση στην υπηρεσία.1.Εξουσιοδότηση χρήστη:Εξουσιοδοτεί ο χρήστης τον εξυπηρετητή του ΕΛΓΑ, να προσπελάσει τα στοιχεία του (ΑΦΜ), που τηρούνται στην ΑΑΔΕ.1.Επιλογή Αίτησης:Για  την εκτύπωση της Βεβαίωσης αποδοχών τακτικού / εποχικού προσωπικού ΕΛΓΑ, θα πρέπει να μεταβεί ο ενδιαφερόμενος στις αιτήσεις.1.Ανάκτηση  Βεβαίωσης αποδοχών τακτικού / εποχικού προσωπικού ΕΛΓΑ  προς εκτύπωση:Μετά τη σχετική επιλογή, ανακτάται η  Βεβαίωση καταβολής αποζημιώσεων φυτικής παραγωγής ΕΛΓΑ, προς εκτύπωση από τον ενδιαφερόμενο.1.Αποτυχία ανάκτησης Βεβαίωσης  αποδοχών τακτικού / εποχικού προσωπικού ΕΛΓΑ προς εκτύπωση:Εμφανίζεται σχετικό μήνυμα αναφορικά με τους λόγους μη ανάκτησης της Βεβαίωσης  αποδοχών τακτικού / εποχικού προσωπικού ΕΛΓΑ.  "
    },
  ],
  temperature=0,
  max_tokens=3755,
  top_p=1,
  frequency_penalty=0,
  presence_penalty=0
)

content = response.choices[0].message.content
print('content = ', content)

contents = content.split('\n\n')

for i, element in enumerate(contents, 1):
    print(f"Paragraph {i}:\n{element}\n")
    sheet.cell(row=i, column=1, value=element)

# Save the Excel file
workbook.save('outputQuestionsAnswers.xlsx')
