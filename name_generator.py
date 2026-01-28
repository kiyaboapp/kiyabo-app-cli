#!/usr/bin/env python3
"""
Name Generator Script
Generates random names with customizable parameters and exports to CSV.
"""

import argparse
import csv
import random
import sys
from faker import Faker
from typing import List, Dict, Set


class NameGenerator:
    def __init__(self, locale: str = 'en_US', count: int = 100, output_file: str = 'names.csv'):
        self.locale = locale
        self.count = count
        self.output_file = output_file
        self.fake = Faker(locale)
        self.generated_names: Set[str] = set()
        
        # Comprehensive Swahili female names from Tanzania and Kenya
        self.swahili_female_names = [
            'Aadila', 'Abiria', 'Abla', 'Adhra', 'Adia', 'Adila', 'Adimu', 'Afaafa', 'Afiya', 'Ahadi',
            'Ainru', 'Ajia', 'Akilah', 'Akina', 'Alama', 'Alamisi', 'Alasiri', 'Aleela', 'Alila', 'Amali',
            'Amani', 'Ambata', 'Amina', 'Aminia', 'Amne', 'Angaza', 'Anisun', 'Arafa', 'Arifa', 'Asatira',
            'Asha', 'Ashura', 'Asia', 'Asubuhi', 'Asya', 'Aza', 'Aziza', 'Azize', 'Bahati', 'Bakari',
            'Barika', 'Bashira', 'Batini', 'Baya', 'Bayo', 'Bayyina', 'Bibi', 'Bimkubwa', 'Binti', 'Bishara',
            'Busara', 'Bushira', 'Chane', 'Chausiku', 'Chiku', 'Chriki', 'Chuki', 'Dabiku', 'Dafina', 'Dalila',
            'Dalili', 'Dhakiya', 'Ducha', 'Duni', 'Eidi', 'Endesha', 'Eshe', 'Etana', 'Fadiya', 'Fahari',
            'Faika', 'Fanaka', 'Fanana', 'Farashuu', 'Farasi', 'Fathiya', 'Fila', 'Firyali', 'Furaha', 'Gamila',
            'Gasira', 'Gheche', 'Goma', 'Habiba', 'Hadiya', 'Haiba', 'Haifa', 'Hakima', 'Hala', 'Halili',
            'Halima', 'Halisi', 'Hamida', 'Hanifa', 'Hanuni', 'Haoriyao', 'Hasa', 'Hashiki', 'Hasina', 'Haya',
            'Hiari', 'Hiba', 'Hususa', 'Ieeda', 'Imani', 'Inaya', 'Ishi', 'Ita', 'Ituri', 'Jaha', 'Jahi',
            'Jamila', 'Jasira', 'Jina', 'Jioni', 'Jirani', 'Johari', 'Jokha', 'Jokla', 'Juwayria', 'Kabisa',
            'Kadija', 'Kaidi', 'Kalere', 'Kamaria', 'Kameke', 'Kamili', 'Kanai', 'Kanoni', 'Karima', 'Kesi',
            'Khadija', 'Kiaga', 'Kibali', 'Kibibi', 'Kiburi', 'Kidawa', 'Kijakazi', 'Kinaya', 'Kioja', 'Kipenzi',
            'Kisasa', 'Kisima', 'Kivuli', 'Koffi', 'Latifah', 'Leta', 'Lila', 'Lulu', 'Madaha', 'Madiha',
            'Mafunda', 'Maisha', 'Maijani', 'Malaika', 'Maliha', 'Malika', 'Mapenzi', 'Marini', 'Marjani', 'Masara',
            'Masika', 'Mchumba', 'Merah', 'Mkiwa', 'Mosi', 'Moza', 'Mpenzi', 'Msaada', 'Msiba', 'Mtama',
            'Mtumwa', 'Mufiida', 'Munira', 'Mvita', 'Mwajuma', 'Mwaka', 'Mwamini', 'Mwanaidi', 'Mwanajuma',
            'Mwanakweli', 'Mwasaa', 'Mwatabu', 'Mwema', 'Nabila', 'Nadhari', 'Nadiya', 'Nadja', 'Nafisika',
            'Naima', 'Najia', 'Najma', 'Neema', 'Nia', 'Niara', 'Njema', 'Nuha', 'Nuru', 'Nyah', 'Paka',
            'Panya', 'Penda', 'Pili', 'Rabuwa', 'Rafiya', 'Raha', 'Ramla', 'Raziya', 'Rehema', 'Ruqayah',
            'Saada', 'Saba', 'Sabiha', 'Sadaka', 'Sadikika', 'Safisha', 'Safiya', 'Saida', 'Sakina', 'Salima',
            'Samiha', 'Samira', 'Sanaa', 'Sanura', 'Sauda', 'Semeni', 'Shahida', 'Shangilia', 'Shani', 'Sharifa',
            'Shawana', 'Shifaa', 'Sikidhani', 'Sijaona', 'Siri', 'Siti', 'Somo', 'Subira', 'Suhaila', 'Sulayma',
            'Surayya', 'Taabu', 'Tabia', 'Tamasha', 'Tatu', 'Tawa', 'Tisa', 'Tisha', 'Tosha', 'Tufaha',
            'Tumaini', 'Uzima', 'Uzuri', 'Wanyika', 'Waseme', 'Winda', 'Yumna', 'Yusra', 'Zaafarani', 'Zahara',
            'Zahina', 'Zainabu', 'Zakiya', 'Zalika', 'Zawadi', 'Zawati', 'Zuri', 'Zulfa', 'Zuwena', 'Zwena',
            'Aisha', 'Fatuma', 'Mariam', 'Khadija', 'Zainab', 'Amina', 'Halima', 'Safiya', 'Mwanahawa', 'Zawadi',
            'Neema', 'Grace', 'Joyce', 'Miriam', 'Afiya', 'Adila', 'Amani', 'Akina', 'Bahiya', 'Dalila',
            'Hasina', 'Imani', 'Jaha', 'Kamili', 'Kesi', 'Lulu', 'Mila', 'Musu', 'Nina', 'Njeri'
        ]
        
        # Comprehensive Swahili male names from Tanzania and Kenya  
        self.swahili_male_names = [
            'Abasi', 'Abdu', 'Abdullahi', 'Abiria', 'Adamu', 'Adin', 'Afla', 'Ajabu', 'Ajali', 'Akida',
            'Alama', 'Alamisi', 'Alasiri', 'Alfajiri', 'Ali', 'Amali', 'Amani', 'Amaziah', 'Amiri', 'Amwa',
            'Anga', 'Angaza', 'Anza', 'Asani', 'Asante', 'Ashon', 'Ashur', 'Ashura', 'Asili', 'Asimwe',
            'Asubuhi', 'Auni', 'Ayubu', 'Aza', 'Azizi', 'Babu', 'Bado', 'Bahari', 'Bakari', 'Balozi',
            'Bamba', 'Baraka', 'Bashiri', 'Binadamu', 'Bomani', 'Bopo', 'Buibui', 'Chane', 'Chui', 'Chuma',
            'Darweshi', 'Daudi', 'Dipili', 'Dogo', 'Elimu', 'Enzi', 'Erevu', 'Fahamu', 'Fakihi', 'Faraji',
            'Farasi', 'Fariji', 'Feruzi', 'Fumo', 'Fupi', 'Haji', 'Haki', 'Hamadi', 'Hamidi', 'Hamisi',
            'Hanisi', 'Hasa', 'Hasani', 'Hassani', 'Heri', 'Hiji', 'Hishima', 'Hurani', 'Husani', 'Ibada',
            'Idi', 'Idili', 'Ijumaa', 'Imamu', 'Imani', 'Ishara', 'Issa', 'Jaali', 'Jabali', 'Jabari',
            'Jabiri', 'Jafari', 'Jahari', 'Jahi', 'Jalali', 'Jamari', 'Jasiri', 'Jela', 'Jelani', 'Jenebi',
            'Jengo', 'Jimoh', 'Joshi', 'Juma', 'Jumaane', 'Jumange', 'Kabona', 'Kafara', 'Kafil', 'Kamau',
            'Kamili', 'Kamisi', 'Kandoro', 'Kanu', 'Kawawe', 'Khalfani', 'Khamisi', 'Khari', 'Kheri', 'Kibwe',
            'Kifimbo', 'Kijana', 'Kijme', 'Kimani', 'Kiongozi', 'Kipawa', 'Kito', 'Kitwana', 'Kombo', 'Kondo',
            'Kongoresi', 'Kukimbia', 'Kumbufa', 'Kumbuka', 'Kwanza', 'Kweli', 'Machupa', 'Majaliwa', 'Mamba',
            'Mansa', 'Marko', 'Maskini', 'Masud', 'Matata', 'Mbita', 'Mbwana', 'Mchawi', 'Mdogo', 'Mfaume',
            'Mhina', 'Mjibu', 'Mosi', 'Moyo', 'Msabaha', 'Msafiri', 'Msemaji', 'Mshangama', 'Mtafiti', 'Mtawa',
            'Mtembei', 'Mtume', 'Musa', 'Mwaka', 'Mwalimu', 'Mwana', 'Mwenye', 'Mwinyi', 'Mwita', 'Mzale',
            'Mzee', 'Mzuri', 'Mzwanza', 'Nassor', 'Nuru', 'Omari', 'Pendo', 'Pili', 'Pupa', 'Radhi',
            'Rafiki', 'Rajabu', 'Rashidi', 'Sadaka', 'Sadiki', 'Sadiku', 'Safi', 'Safwani', 'Saka', 'Salehe',
            'Salin', 'Sefu', 'Shaaboni', 'Shamba', 'Shangwe', 'Shani', 'Shibe', 'Shibisha', 'Shida', 'Shomari',
            'Sifa', 'Simba', 'Siwatu', 'Siwazuri', 'Songoro', 'Sudi', 'Suhuba', 'Sulubu', 'Tafiti', 'Tajiri',
            'Tawfiki', 'Tayari', 'Tembo', 'Tendaji', 'Tiifu', 'Tisa', 'Tumaini', 'Ubora', 'Ubwa', 'Ufanisi',
            'Utende', 'Utendi', 'Uzuri', 'Vuai', 'Waziri', 'Yakubu', 'Yohana', 'Yusufu', 'Zahir', 'Zahur',
            'Zakia', 'Zakwani', 'Zwadi', 'Zuberi', 'Azizi', 'Ahmed', 'Abdalla', 'Azaan', 'Bahati', 'Baraka',
            'Daudi', 'Faraji', 'Hamidi', 'Hasani', 'Jirani', 'Kabili', 'Kanai', 'Kenyada', 'Nafasi', 'Penda',
            'Radhi', 'Sadiki', 'Salimu', 'Sefu', 'Tabari', 'Tumaini', 'Yahya', 'Zahur', 'Zahoor', 'Adili',
            'Akida', 'Asani', 'Ayubu', 'Bakari', 'Balozi', 'Chacha', 'Damu', 'Haki', 'Hami', 'Jabali',
            'Jel', 'Jelani', 'Kenyatta', 'Kito', 'Rafiki', 'Sadeeki', 'Sahel', 'Simba', 'Taha', 'Yazeed',
            'Yazid', 'Zahour', 'Zubery', 'Hassan', 'Mohamed', 'Ibrahim', 'Abdullah', 'Yusuf', 'Omar', 'Juma'
        ]
        
        # Comprehensive Swahili and Tanzanian surnames
        self.swahili_surnames = [
            'Abdala', 'Abdalah', 'Abdalla', 'Abdallah', 'Abdul', 'Abdulaziz', 'Abel', 'Ackson', 'Adam',
            'Akukweti', 'Alex', 'Alfred', 'Ali', 'Ally', 'Alouce', 'Alphonce', 'Amiri', 'Amos', 'Amour',
            'Andrea', 'Andrew', 'Anthony', 'Athuman', 'Athumani', 'Bakari', 'Bilal', 'Boniphace', 'Chacha',
            'Charles', 'Chaula', 'Chegeni', 'Christopher', 'Cosmas', 'Daniel', 'Daud', 'Daudi', 'David',
            'Deus', 'Edward', 'Elias', 'Emanuel', 'Emmanuel', 'Ernest', 'Ezekiel', 'Faustine', 'Francis',
            'Frank', 'Gabriel', 'George', 'Ghailani', 'Godfrey', 'Haji', 'Hamad', 'Hamadi', 'Hamis', 'Hamisi',
            'Hamza', 'Haruna', 'Hassan', 'Hassani', 'Haule', 'Hussein', 'Ibrahim', 'Ibrahimu', 'Iddi', 'Ismail',
            'Issa', 'Jackson', 'Jacob', 'Jamal', 'James', 'January', 'Japhet', 'John', 'Jonas', 'Joseph',
            'Josephat', 'Julius', 'Juma', 'Jumanne', 'Jumbe', 'Kalinga', 'Kapinga', 'Karamagi', 'Karoli',
            'Karume', 'Kassim', 'Kassum', 'Kawawa', 'Kayombo', 'Kessy', 'Kevela', 'Khamis', 'Kibona', 'Kikwete',
            'Kimaro', 'Kinana', 'Kinasha', 'Kinyonga', 'Kitali', 'Kolimba', 'Komba', 'Kombo', 'Kulwa', 'Kyando',
            'Laizer', 'Lameck', 'Laurent', 'Lazaro', 'Lema', 'Leonard', 'Lucas', 'Lyimo', 'Mabula', 'Maganga',
            'Magesa', 'Magufuli', 'Mahenge', 'Makamba', 'Makame', 'Makoye', 'Malecela', 'Malima', 'Mangula',
            'Mapunda', 'Marco', 'Martin', 'Marwa', 'Masanja', 'Mashaka', 'Masoud', 'Massawe', 'Mathias',
            'Mbilinyi', 'Mbise', 'Mbwambo', 'Mbwana', 'Mdee', 'Meghji', 'Mfaki', 'Mgaya', 'Mgeni', 'Mhagama',
            'Mhando', 'Mhina', 'Michael', 'Minja', 'Mkapa', 'Mlay', 'Mohamed', 'Mohamedi', 'Mollem', 'Mollel',
            'Moses', 'Mosha', 'Moshi', 'Mrema', 'Mrisho', 'Msangi', 'Mshana', 'Msekela', 'Msekwa', 'Msigwa',
            'Msuya', 'Mukama', 'Musa', 'Mushi', 'Mussa', 'Mwambi', 'Mwinuka', 'Mwinyi', 'Mwita', 'Nassor',
            'Nassoro', 'Nchimbi', 'Ndunguru', 'Ngasongwa', 'Ngombale', 'Ngonyani', 'Ngowi', 'Nnauye', 'Nyerere',
            'Nyoni', 'Omar', 'Omari', 'Omary', 'Ongala', 'Othman', 'Paschal', 'Patrick', 'Paul', 'Paulo',
            'Peter', 'Petro', 'Philipo', 'Pius', 'Rajabu', 'Ramadhan', 'Ramadhani', 'Raphael', 'Rashid', 'Rashidi',
            'Raza', 'Richard', 'Robert', 'Said', 'Saidi', 'Salehe', 'Salim', 'Salimu', 'Salla', 'Salum',
            'Salumu', 'Samson', 'Samwel', 'Sanga', 'Seif', 'Seleman', 'Selemani', 'Shaban', 'Shabani', 'Shayo',
            'Shija', 'Shirima', 'Simba', 'Simon', 'Sokoine', 'Stephano', 'Stephen', 'Suleiman', 'Sumari', 'Sumaye',
            'Swai', 'Tarimo', 'Temba', 'Temu', 'Thomas', 'Tuweni', 'Wakil', 'Wambura', 'Wapakhabulo', 'Warioba',
            'Waziri', 'William', 'Wilson', 'Yahaya', 'Yohana', 'Yusuph', 'Zacharia', 'Zawose', 'Zuberi',
            'Mwangi', 'Karanja', 'Njoroge', 'Kimani', 'Muriuki', 'Wanjiru', 'Wairimu', 'Nyambura', 'Gatonye',
            'Thiongo', 'Mbugua', 'Kamau', 'Maina', 'Kinyua', 'Muriithi', 'Waweru', 'Kiragu', 'Macharia',
            'Njenga', 'Mwangi', 'Kariuki', 'Njoku', 'Okonkwo', 'Okafor', 'Eze', 'Okoro', 'Iweala', 'Adebayo',
            'Bakare', 'Ogunleye', 'Adewale', 'Oladipo', 'Osunde', 'Ezeji', 'Nwosu', 'Okonkwo', 'Obi', 'Anyanwu'
        ]
    
    def generate_name(self, sex: str = None) -> Dict[str, str]:
        """Generate a single name with first, middle, surname, and sex."""
        if sex is None:
            sex = random.choice(['M', 'F'])
        
        # Use Swahili names for East African locales (60% chance for better representation)
        use_swahili = random.random() < 0.6
        
        if use_swahili and self.locale in ['sw', 'ke', 'tz']:
            if sex == 'M':
                first_name = random.choice(self.swahili_male_names)
                middle_name = random.choice(self.swahili_male_names)
                surname = random.choice(self.swahili_surnames)
            else:
                first_name = random.choice(self.swahili_female_names)
                middle_name = random.choice(self.swahili_male_names)  # Middle names often masculine
                surname = random.choice(self.swahili_surnames)
        else:
            # Use Faker for standard names
            if sex == 'M':
                first_name = self.fake.first_name_male()
                middle_name = self.fake.first_name_male()
                surname = self.fake.last_name()
            else:
                first_name = self.fake.first_name_female()
                middle_name = self.fake.first_name_male()  # Middle names are masculine
                surname = self.fake.last_name()
        
        name_record = {
            'first_name': first_name,
            'middle_name': middle_name,
            'surname': surname,
            'sex': sex
        }
        
        return name_record
    
    def generate_unique_names(self, sex: str = None) -> List[Dict[str, str]]:
        """Generate unique names, avoiding duplicates."""
        names = []
        attempts = 0
        max_attempts = self.count * 10  # Prevent infinite loops
        
        while len(names) < self.count and attempts < max_attempts:
            name = self.generate_name(sex)
            
            # Create a unique identifier for the name
            name_key = f"{name['first_name']}_{name['middle_name']}_{name['surname']}_{name['sex']}"
            
            if name_key not in self.generated_names:
                self.generated_names.add(name_key)
                names.append(name)
            
            attempts += 1
        
        if len(names) < self.count:
            print(f"Warning: Could only generate {len(names)} unique names out of {self.count} requested.", file=sys.stderr)
        
        return names
    
    def export_to_csv(self, names: List[Dict[str, str]]) -> None:
        """Export names to CSV file."""
        fieldnames = ['first_name', 'middle_name', 'surname', 'sex']
        
        with open(self.output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(names)
        
        print(f"Successfully exported {len(names)} names to {self.output_file}")


def main():
    parser = argparse.ArgumentParser(description='Generate random names with customizable parameters')
    
    parser.add_argument('-c', '--count', type=int, default=100,
                        help='Number of names to generate (default: 100)')
    
    parser.add_argument('-l', '--locale', type=str, default='sw',
                        choices=['en_US', 'sw', 'ke', 'tz'],
                        help='Locale for name generation (en_US, sw, ke, tz) (default: sw)')
    
    parser.add_argument('-s', '--sex', type=str, choices=['M', 'F'],
                        help='Sex for generated names (M or F). If not specified, names will be random')
    
    parser.add_argument('-o', '--output', type=str, default='names.csv',
                        help='Output CSV file name (default: names.csv)')
    
    args = parser.parse_args()
    
    # Map locale choices to proper Faker locales
    locale_mapping = {
        'en_US': 'en_US',
        'sw': 'sw',  # Swahili (if available)
        'ke': 'en_KE',  # Kenya English
        'tz': 'en_TZ'   # Tanzania English
    }
    
    locale = locale_mapping.get(args.locale, 'en_US')
    
    # Create generator and generate names
    generator = NameGenerator(locale=locale, count=args.count, output_file=args.output)
    
    print(f"Generating {args.count} names with locale '{args.locale}'...")
    if args.sex:
        print(f"Sex: {args.sex}")
    else:
        print("Sex: Random")
    
    names = generator.generate_unique_names(args.sex)
    generator.export_to_csv(names)
    
    # Display sample
    print("\nSample generated names:")
    for i, name in enumerate(names[:5], 1):
        print(f"{i}. {name['first_name']} {name['middle_name']} {name['surname']} ({name['sex']})")


if __name__ == '__main__':
    main()
