// ============================================================
// APLIKASI FOMO - Skoring RIASEC Holland (Standar Excel Holland Test)
// Backend : Google Apps Script (Code.gs)
// Database: Google Sheets (dibuat OTOMATIS)
// ============================================================

// Nama sheet di Spreadsheet
var SHEET_TOKEN     = 'Token';
var SHEET_BIODATA   = 'Biodata';
var SHEET_SOAL      = 'Soal';
var SHEET_JAWABAN   = 'Jawaban';
var SHEET_HASIL     = 'Hasil';
var SHEET_ADMIN     = 'Admin';

// Nama file Spreadsheet yang akan dibuat otomatis
var SPREADSHEET_NAME = 'FOMO Database';

// ---- MAPPING BIDANG RIASEC Holland ----
var BIDANG_MAP = {
  'R': 'Realistik',    // Realistic    - Teknik, Mekanik, Praktis
  'I': 'Investigatif', // Investigative- Sains, Penelitian, Analitis
  'A': 'Artistik',     // Artistic     - Seni, Kreatif, Ekspresi
  'S': 'Sosial',       // Social       - Menolong, Mengajar, Relasi
  'E': 'Enterprising', // Enterprising - Kepemimpinan, Bisnis, Persuasi
  'C': 'Konvensional'  // Conventional - Administrasi, Prosedur, Data
};

var BIDANG_INGGRIS = {
  'R': 'Realistic',
  'I': 'Investigative',
  'A': 'Artistik',
  'S': 'Social',
  'E': 'Enterprising',
  'C': 'Conventional'
};

// ============================================================
// DATA REFERENSI HOLLAND (Sesuai Sheet Referensi Excel)
// ============================================================
var REFERENSI_RIASEC = {"single": {"R": {"deskripsi": "Profil ini menyukai pekerjaan yang mencakup masalah dan jawaban praktis dan langsung. Menyukai kegiatan yang melibatkan keterampilan motorik, peralatan, mesin Seringkali orang dengan minat Realistis tidak menyukai karier yang melibatkan dokumen atau bekerja sama dengan orang lain.", "profesi": "Mekanik, Insinyur, pengawas bangunan, Petani /Peternak Modern, Atlet, Operator, Polisi, Pemadam Kebakaran, Koki"}, "I": {"deskripsi": "Profil ini menyukai pekerjaan yang berkaitan dengan ide dan pemikiran. memilih kegiatan penyelidikan yang observasional, simbolis, sistematis dan kreatif atas fenomena fisik, biologis dan kultural dalam usahanya untuk mengerti dan mengendalikan fenomena tersebut.  Mereka berprestasi terutama dalam bidang akademik dan ilmiah, serta biasanya kurang menyukai aktivitas fisik atau memimpin orang.", "profesi": "Ahli antropologi, astronomi, biologi, botani, kimia, editor penerbitan ilmiah, geologi, pekerjaan penelitian, dan pekerjaan lain yang sejenis"}, "A": {"deskripsi": "Profil ini menyukai pekerjaan dan aktivitas yang bebas, tidak sistematis, ambisius dalam usahanya memanipulasi material fisik, verbal dan humanistic untuk tujuan menciptakan bentuk-bentuk atau produk-produk seni. Kecenderungan ini mengantarkannya memiliki kompetensi artistik, bahasa, musik, drama, mengarang. Tetapi mengurangi kompetensi dibidang klerikal dan bisnis. ", "profesi": "Desainer, Penulis, Musisi, Arsitek, Wartawan, Penari, Translator, Artis dan pekerjaan lain yang sejenis."}, "S": {"deskripsi": "Profil ini menyukai aktivitas yang melibatkan berelasi, berkomunikasi dan mengajar orang. Mereka dilambangkan dengan ketrampilan sosial dan kebutuhan untuk interaksi sosial. Pekerjaannya sering melibatkan membantu atau memberikan layanan kepada orang lain. Tetapi mengurangi kompetensi di bidang teknik dan manual.  ", "profesi": "Guru, Terapis, Tour Guide, Perawat, Hakim, Konselor, Psikolog, Sejarawan, misionari, Pekerja Sosial, Ahli Patolog Bicara dan pekerjaan lain yang sejenis."}, "E": {"deskripsi": "Profil ini menghasilkan preferensi yang mendorong individu untuk menikmati memanipulasi kegiatan orang lain dalam usahanya mencapai tujuan organisasi atau keuntungan ekonomi. Kecenderungan ini akan mengarah pada kompetensi kepemimpinan, hubungan timbal balik pribadi dan persuasif, tugas yang menitikberatkan kepada kemampuan verbal yang digunakan untuk mengarahkan dan mempengaruhi orang lain. Tetapi mengurangi kompetensi ilmiah.", "profesi": "Pengacara, Pelobi, Politikus, Pegawai humas, Pelobi, Manajer penjualan, Staf penjualan, perwakilan dagang, dan pekerjaan lain yang sejenis."}, "C": {"deskripsi": "Profil ini menghasilkan preferensi yang membawa individu untuk menyenangi aktivitas yang teratur dan sistematis seperti misalnya mengarsipkan sesuatu, mereproduksi material, mencatat data, mengolah data, menyukai pekerjaan yang mengikuti prosedur dan rutinitas yang ditetapkan. Kecenderungan ini me-ngantarkan individu memiliki kompetensi klerikal, menghitung / kompetensi Operasional bisnis. Tapi mengurangi kompetensi artistik.", "profesi": "Akuntan, Aktuaris, Ahli Statistik, Analis Keuangan, Operator komputer, Pustakawan, Programer Bisnis, Pengarsip, dan pekerjaan yang berhubungan dengan administrasi."}}, "konsistensi": {"RI": "Tinggi", "RC": "Tinggi", "IR": "Tinggi", "IA": "Tinggi", "AI": "Tinggi", "AS": "Tinggi", "SA": "Tinggi", "SE": "Tinggi", "ES": "Tinggi", "CE": "Tinggi", "EC": "Tinggi", "CR": "Tinggi", "RA": "Sedang", "RE": "Sedang", "IS": "Sedang", "IC": "Sedang", "AR": "Sedang", "AE": "Sedang", "SI": "Sedang", "SC": "Sedang", "EA": "Sedang", "ER": "Sedang", "CS": "Sedang", "CI": "Sedang", "RS": "Rendah", "IE": "Rendah", "AC": "Rendah", "SR": "Rendah", "EI": "Rendah", "CA": "Rendah"}, "combo2": {"AC": {"profesi": "Graphic Designer, Proofreader, Web/Mobile Design, UI Designer & Visual Designer,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan yang kreatif dan berorientasi pada proses. Jenis pekerjaan ini membutuhkan proses pemikiran artistik dan cara kerja yang sehat, andal, dan terstruktur."}, "AS": {"profesi": "Guru, Translator, Nanny, Career Advisor & Interpreter,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan yang kreatif dan berorientasi pada manusia. Jenis pekerjaan ini membutuhkan pikiran artistik dan sosial."}, "EC": {"profesi": "General manager, Air traffic controller, Branch Manager, Fundraiser, Recruiter & Shop Assistant,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan yang giat dan berbasis proses. Jenis pekerjaan ini membutuhkan banyak tanggung jawab pribadi dan cara kerja yang sehat, andal, dan terstruktur."}, "IA": {"profesi": "Architect, Technical Writer, Animator & Desktop Publisher,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan yang berorientasi pada penelitian dan kreatif. Ini adalah jenis pekerjaan yang membutuhkan proses pemikiran analitis dan keingintahuan ilmiah. Pada saat yang sama, itu membutuhkan pola pikir yang unik dan artistik."}, "IC": {"profesi": "Pengembang perangkat lunak , UX Designer, Paralegal, Data Analyst & Quality Controller,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan penelitian dan pendekatan berbasis proses. Jenis pekerjaan ini membutuhkan proses pemikiran analitis dan cara kerja yang sehat, andal, dan terstruktur."}, "IE": {"profesi": "Consultant,Detektif, Solicitor, Market Researcher, Business Controller & Online marketer,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan yang berorientasi pada penelitian dan giat. Jenis pekerjaan ini membutuhkan proses pemikiran analitis, semangat giat dan rasa kepemimpinan dan tanggung jawab."}, "IS": {"profesi": "Dokter, Physics Teacher, Nurse, Physiotherapist & Coach,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan penelitian dan pendekatan yang berorientasi pada manusia. Jenis pekerjaan ini membutuhkan proses pemikiran analitis dan pola pikir sosial."}, "RA": {"profesi": "Photographer, Tailor, Artist, Pastry Chef & Furniture Designer,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali mencakup aktivitas yang melibatkan pendekatan fisik dan kreatif. Jenis pekerjaan ini membutuhkan tenaga kerja manual di satu sisi dan pola pikir imajinatif di sisi lain."}, "RE": {"profesi": "Koki, Engineering Manager, Police Officer, Process Operator & Taxi Driver,", "deskripsi": "profesi yang terkait dengan profil ini sering kali melibatkan pendekatan fisik dan giat. Jenis pekerjaan ini membutuhkan tenaga kerja manual, mengambil inisiatif dan rasa tanggung jawab."}, "AE": {"profesi": "Design Interior, Illustrator, Entrepreneur, Journalist, Producer, Art director, Copywriter & Actor,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan yang kreatif dan giat. Jenis pekerjaan ini membutuhkan pikiran artistik, rasa tanggung jawab atau kadang-kadang bahkan kepemimpinan."}, "RI": {"profesi": "Network Engineer, Mechanical Engineer, Automotive Engineer, Electronic Enginer, Building Insulation Installer, Pilot,  Laboratory Technician, Fruit and Vegetable Factory Worker, Product Assembler,", "deskripsi": "Profil ini terkait dengan profesi yang melibatkan banyak pekerjaan manual, pekerjaan di luar ruangan atau berhubungan dengan alat, mesin, benda, binatang atau tanaman. Pada saat yang sama, jenis profesi ini membutuhkan pemikiran yang cermat dan pendekatan yang lebih analitis."}, "RS": {"profesi": "Asisten perawatan, Fire fighter, Masseur, Sports Coach & Veterinary Assistant,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan fisik dan berorientasi pada manusia. Jenis pekerjaan ini membutuhkan tenaga kerja manual dan membantu orang lain."}, "SC": {"profesi": "Librarian, Telephone Operator, Medical Assistant & Teaching Assistant,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan yang berorientasi pada orang dan berbasis proses. Jenis pekerjaan ini membutuhkan bantuan orang lain atas dasar cara kerja yang baik, andal, dan terstruktur."}, "RC": {"profesi": "Lorry Driver, Machine Operator, Security Guard, Plumber, Cleaner, Warehouse operative & Service technician,", "deskripsi": "Profesi yang terkait  Profil ini terkait dengan profesi yang banyak melibatkan tenaga fisik, manual dan outdoor yang melibatkan mesin. Pada saat yang sama, profesi semacam ini membutuhkan pendekatan yang cermat, teratur, dan terstruktur."}, "SE": {"profesi": "HR Manager, Account Manager, Customer Service Advisor & Waiter / Waitress,", "deskripsi": "Profesi yang terkait dengan profil ini sering kali melibatkan pendekatan yang berorientasi pada orang dan giat. Jenis pekerjaan ini membutuhkan membantu orang lain di satu sisi dan mengambil kepemimpinan dan tanggung jawab di sisi lain."}}, "profesi_kombinasi": {"AE": "Actors, Art Directors, Interior Designers, Video Game Designers,", "AEC": "Editors,", "AEI": "Film and Video Editors, News Analysts, Reporters, and Journalists,", "AER": "Commercial and Industrial Designers, Fashion Designers, Floral Designers, Merchandise Displayers and Window Trimmers, Models,", "AES": "Broadcast Announcers and Radio Disc Jockeys, Hairdressers, Hairstylists, and Cosmetologists, Music Directors and Composers, Musicians and Singers,", "AI": "Architects, Except Landscape and Naval, Poets, Lyricists and Creative Writers, Special Effects Artists and Animators,", "AIC": "Desktop Publishers, Technical Writers,", "AIR": "Landscape Architects,", "AR": "Dancers, Fine Artists, Including Painters, Sculptors, and Illustrators, Makeup Artists, Theatrical and Performance, Photographers, Set and Exhibit Designers,", "ARC": "Cooks, Private Household,", "ARE": "Craft Artists, Graphic Designers,", "AS": "Interpreters and Translators,", "ASE": "Choreographers,", "C": "Document Management Specialists,", "CA": "Proofreaders and Copy Markers,", "CE": "Bill and Account Collectors, Bookkeeping, Accounting, and Auditing Clerks, Brokerage Clerks, Cashiers, Compensation, Benefits, and Job Analysis Specialists, Correspondence Clerks, Cost Estimators, Counter and Rental Clerks, Court Reporters and Simultaneous Captioners, Credit Analysts, Energy Auditors, Executive Secretaries and Executive Administrative Assistants, Freight Forwarders, Legal Secretaries and Administrative Assistants, Loan Interviewers and Clerks, Payroll and Timekeeping Clerks, Procurement Clerks, Production, Planning, and Expediting Clerks, Purchasing Agents, Except Wholesale, Retail, and Farm Products, Regulatory Affairs Specialists, Sales Representatives, Wholesale and Manufacturing, Except Technical and Scientific Products, Secretaries and Administrative Assistants, Except Legal, Medical, and Executive, Tax Examiners and Collectors, and Revenue Agents, Tax Preparers, Tellers, Treasurers and Controllers, Word Processors and Typists,", "CEI": "Accountants and Auditors, Budget Analysts, Claims Adjusters, Examiners, and Investigators, Insurance Underwriters, Judicial Law Clerks, Logistics Analysts, Web Administrators,", "CER": "Cargo and Freight Agents, Compliance Managers, Customs and Border Protection Officers, Gambling and Sports Book Writers and Runners, Gambling Cage Workers, Gambling Change Persons and Booth Cashiers, Gambling Dealers, Government Property Inspectors and Investigators, Office Clerks, General, Title Examiners, Abstractors, and Searchers,", "CES": "Billing and Posting Clerks, Court, Municipal, and License Clerks, Credit Authorizers, Checkers, and Clerks, Hotel, Motel, and Resort Desk Clerks, Human Resources Assistants, Except Payroll and Timekeeping, Insurance Claims and Policy Processing Clerks, Interviewers, Except Eligibility and Loan, Loan Officers, New Accounts Clerks, Order Clerks, Receptionists and Information Clerks, Reservation and Transportation Ticket Agents and Travel Clerks, Switchboard Operators, Including Answering Service,", "CI": "Archivists, Clinical Data Managers, Database Administrators, Social Science Research Assistants, Statistical Assistants, Statisticians,", "CIE": "Actuaries, Paralegals and Legal Assistants,", "CIR": "Computer Numerically Controlled Tool Programmers, Environmental Compliance Inspectors, Information Security Analysts, Quality Control Analysts, Web Developers,", "CR": "Costume Attendants, Inspectors, Testers, Sorters, Samplers, and Weighers, Mail Clerks and Mail Machine Operators, Except Postal Service, Medical Transcriptionists, Meter Readers, Utilities, Occupational Health and Safety Technicians, Pharmacy Aides, Pharmacy Technicians, Photographic Process Workers and Processing Machine Operators, Postal Service Clerks, Postal Service Mail Carriers, Postal Service Mail Sorters, Processors, and Processing Machine Operators, Weighers, Measurers, Checkers, and Samplers, Recordkeeping,", "CRE": "Data Entry Keyers, Dispatchers, Except Police, Fire, and Ambulance, File Clerks, Insurance Appraisers, Auto Damage, Public Safety Telecommunicators, Shipping, Receiving, and Inventory Clerks, Stockers and Order Fillers,", "CRI": "Ophthalmic Medical Technologists, Phlebotomists, Police Identification and Records Officers, Surveying and Mapping Technicians,", "CRS": "Dental Assistants, Library Assistants, Clerical, Orderlies,", "CS": "Medical Secretaries and Administrative Assistants,", "CSE": "Library Technicians, Social and Human Service Assistants,", "CSR": "Librarians and Media Collections Specialists, Ophthalmic Medical Technicians, Surgical Assistants, Telephone Operators,", "EA": "Talent Directors,", "EAC": "Advertising and Promotions Managers, Writers and Authors,", "EAS": "Producers and Directors, Public Relations Specialists,", "EC": "Administrative Services Managers, Air Traffic Controllers, Chief Executives, Curators, Customs Brokers, Door-to-Door Sales Workers, News and Street Vendors, and Related Workers, Financial Examiners, Gambling Managers, Geothermal Production Managers, Hydroelectric Production Managers, Industrial Production Managers, Information Technology Project Managers, Investment Fund Managers, Logisticians, Loss Prevention Managers, Marketing Managers, Private Detectives and Investigators, Property, Real Estate, and Community Association Managers, Purchasing Managers, Real Estate Brokers, Real Estate Sales Agents, Recycling Coordinators, Regulatory Affairs Managers, Retail Loss Prevention Specialists, Retail Salespersons, Sales Managers, Sales Representatives, Wholesale and Manufacturing, Technical and Scientific Products, Supply Chain Managers, Telemarketers, Travel Agents, Travel Guides, Wholesale and Retail Buyers, Except Farm Products,", "ECA": "Advertising Sales Agents, Fundraisers, Media Programming Directors,", "ECI": "Chief Sustainability Officers, Computer and Information Systems Managers, Wind Energy Development Managers,", "ECR": "Airfield Operations Specialists, Amusement and Recreation Attendants, Appraisers and Assessors of Real Estate, Baristas, Bartenders, Biofuels Production Managers, Biomass Power Plant Managers, Buyers and Purchasing Agents, Farm Products, Demonstrators and Product Promoters, First-Line Supervisors of Correctional Officers, First-Line Supervisors of Food Preparation and Serving Workers, First-Line Supervisors of Gambling Services Workers, First-Line Supervisors of Housekeeping and Janitorial Workers, First-Line Supervisors of Material-Moving Machine and Vehicle Operators, First-Line Supervisors of Mechanics, Installers, and Repairers, Food Service Managers, Online Merchants, Opticians, Dispensing, Parts Salespersons, Quality Control Systems Managers, Solar Sales Representatives and Assessors, Transportation, Storage, and Distribution Managers, Wind Energy Operations Managers,", "ECS": "Compensation and Benefits Managers, Compliance Officers, Education Administrators, Postsecondary, Financial Managers, First-Line Supervisors of Non-Retail Sales Workers, First-Line Supervisors of Office and Administrative Support Workers, First-Line Supervisors of Personal Service Workers, First-Line Supervisors of Retail Sales Workers, Funeral Home Managers, General and Operations Managers, Human Resources Specialists, Insurance Sales Agents, Labor Relations Specialists, Lodging Managers, Medical and Health Services Managers, Meeting, Convention, and Event Planners, Personal Financial Advisors, Postmasters and Mail Superintendents, Securities, Commodities, and Financial Services Sales Agents, Spa Managers,", "EI": "Biofuels/Biodiesel Technology and Product Development Managers, Lawyers, Natural Sciences Managers,", "EIA": "Sustainability Specialists,", "EIC": "Brownfield Redevelopment Specialists and Site Managers, Business Continuity Planners, Clinical Research Coordinators, Detectives and Criminal Investigators, Fraud Examiners, Investigators and Analysts, Search Marketing Strategists,", "EIS": "Administrative Law Judges, Adjudicators, and Hearing Officers,", "ER": "Aircraft Cargo Handling Supervisors, Solar Energy Installation Managers,", "ERA": "Chefs and Head Cooks,", "ERC": "Construction Managers, Farmers, Ranchers, and Other Agricultural Managers, First-Line Supervisors of Construction Trades and Extraction Workers, First-Line Supervisors of Farming, Fishing, and Forestry Workers, First-Line Supervisors of Helpers, Laborers, and Material Movers, Hand, First-Line Supervisors of Landscaping, Lawn Service, and Groundskeeping Workers, First-Line Supervisors of Production and Operating Workers, Media Technical Directors/Managers, Railroad Conductors and Yardmasters,", "ERI": "Architectural and Engineering Managers, Sales Engineers,", "ERS": "Farm Labor Contractors, First-Line Supervisors of Firefighting and Prevention Workers, Police and Sheriff's Patrol Officers,", "ES": "Agents and Business Managers of Artists, Performers, and Athletes, Fitness and Wellness Coordinators, Hosts and Hostesses, Restaurant, Lounge, and Coffee Shop, Judges, Magistrate Judges, and Magistrates, Legislators, Social and Community Service Managers, Training and Development Managers,", "ESC": "Clinical Nurse Specialists, Credit Counselors, Customer Service Representatives, Directors, Religious Activities and Education, Education Administrators, Kindergarten through Secondary, First-Line Supervisors of Police and Detectives, Flight Attendants, Human Resources Managers, Morticians, Undertakers, and Funeral Arrangers,", "ESR": "Skincare Specialists,", "I": "Historians,", "IAR": "Anthropologists and Archeologists, Astronomers, Biochemists and Biophysicists, Geneticists,", "IAS": "Political Scientists, Sociologists,", "IC": "Biostatisticians, Computer Programmers, Data Warehousing Specialists, Financial Quantitative Analysts, Occupational Health and Safety Specialists,", "ICA": "Mathematicians,", "ICE": "Computer Network Architects, Database Architects, Economists, Industrial Engineers, Intelligence Analysts, Operations Research Analysts, Survey Researchers,", "ICR": "Bioinformatics Scientists, Computer Systems Analysts, Geodetic Surveyors, Logistics Engineers, Software Quality Assurance Analysts and Testers, Transportation Planners,", "ICS": "Pharmacists,", "IE": "Climate Change Policy Analysts, Industrial Ecologists,", "IEA": "Industrial-Organizational Psychologists, Urban and Regional Planners,", "IEC": "Business Intelligence Analysts, Environmental Economists, Management Analysts, Market Research Analysts and Marketing Specialists, Water Resource Specialists,", "IR": "Aerospace Engineers, Animal Scientists, Atmospheric and Space Scientists, Bioengineers and Biomedical Engineers, Biologists, Chemical Engineers, Cytotechnologists, Electrical Engineers, Electronics Engineers, Except Computer, Energy Engineers, Except Wind and Solar, Geoscientists, Except Hydrologists and Geographers, Human Factors Engineers and Ergonomists, Hydrologists, Materials Scientists, Microbiologists, Physicians, Pathologists, Physicists, Prosthodontists, Soil and Plant Scientists, Veterinarians, Zoologists and Wildlife Biologists,", "IRA": "Geographers, Marine Engineers and Naval Architects, Medical Scientists, Except Epidemiologists, Molecular and Cellular Biologists,", "IRC": "Bioinformatics Technicians, Chemical Technicians, Chemists, Computer and Information Research Scientists, Computer Hardware Engineers, Computer Systems Engineers/Architects, Coroners, Cytogenetic Technologists, Environmental Engineers, Environmental Science and Protection Technicians, Including Health, Environmental Scientists and Specialists, Including Health, Food Scientists and Technologists, Forensic Science Technicians, Geographic Information Systems Technologists and Technicians, Health and Safety Engineers, Except Mining Safety Engineers and Inspectors, Industrial Engineering Technologists and Technicians, Mechanical Engineers, Mechatronics Engineers, Medical and Clinical Laboratory Technologists, Microsystems Engineers, Network and Computer Systems Administrators, Nuclear Engineers, Petroleum Engineers, Photonics Engineers, Robotics Engineers, Validation Engineers,", "IRE": "Agricultural Engineers, Environmental Restoration Planners, Fire-Prevention and Protection Engineers, Materials Engineers, Mining and Geological Engineers, Including Mining Safety Engineers, Nanosystems Engineers, Water/Wastewater Engineers,", "IRS": "Anesthesiologists, Dentists, General, Nuclear Medicine Technologists, Nurse Anesthetists, Orthodontists, Radiologists,", "IS": "Audiologists, Dietitians and Nutritionists, Epidemiologists, Family Medicine Physicians, Naturopathic Physicians, Pediatricians, General, School Psychologists,", "ISA": "Clinical and Counseling Psychologists, Psychiatrists,", "ISR": "Allergists and Immunologists, Dermatologists, Diagnostic Medical Sonographers, General Internal Medicine Physicians, Neurologists, Obstetricians and Gynecologists, Ophthalmologists, Except Pediatric, Optometrists, Orthoptists, Physical Medicine and Rehabilitation Physicians, Podiatrists, Sports Medicine Physicians, Urologists,", "R": "Agricultural Equipment Operators, Animal Trainers, Automotive Body and Related Repairers, Commercial Divers, Dredge Operators, Excavating and Loading Machine and Dragline Operators, Surface Mining, Extruding and Forming Machine Setters, Operators, and Tenders, Synthetic and Glass Fibers, Farmworkers, Farm, Ranch, and Aquacultural Animals, Floor Sanders and Finishers, Grinding, Lapping, Polishing, and Buffing Machine Tool Setters, Operators, and Tenders, Metal and Plastic, Helpers--Brickmasons, Blockmasons, Stonemasons, and Tile and Marble Setters, Helpers--Pipelayers, Plumbers, Pipefitters, and Steamfitters, Insulation Workers, Floor, Ceiling, and Wall, Laborers and Freight, Stock, and Material Movers, Hand, Meat, Poultry, and Fish Cutters and Trimmers, Mechanical Door Repairers, Motorcycle Mechanics, Multiple Machine Tool Setters, Operators, and Tenders, Metal and Plastic, Painting, Coating, and Decorating Workers, Pipelayers, Plasterers and Stucco Masons, Rail-Track Laying and Maintenance Equipment Operators, Roustabouts, Oil and Gas, Septic Tank Servicers and Sewer Pipe Cleaners, Sheet Metal Workers, Stonemasons, Tapers, Terrazzo Workers and Finishers, Textile Bleaching and Dyeing Machine Operators and Tenders, Tree Trimmers and Pruners, Upholsterers,", "RA": "Camera Operators, Television, Video, and Film, Furniture Finishers, Museum Technicians and Conservators, Potters, Manufacturing, Sound Engineering Technicians, Stone Cutters and Carvers, Manufacturing,", "RAC": "Fabric and Apparel Patternmakers, Jewelers and Precious Stone and Metal Workers, Model Makers, Wood, Sewers, Hand,", "RAE": "Tailors, Dressmakers, and Custom Sewers,", "RAI": "Musical Instrument Repairers and Tuners,", "RC": "Adhesive Bonding Machine Operators and Tenders, Aircraft Structure, Surfaces, Rigging, and Systems Assemblers, Animal Caretakers, Audiovisual Equipment Installers and Repairers, Automotive and Watercraft Service Attendants, Baggage Porters and Bellhops, Bakers, Boilermakers, Bus and Truck Mechanics and Diesel Engine Specialists, Cabinetmakers and Bench Carpenters, Chemical Equipment Operators and Tenders, Chemical Plant and System Operators, Cleaners of Vehicles and Equipment, Cleaning, Washing, and Metal Pickling Equipment Operators and Tenders, Coin, Vending, and Amusement Machine Servicers and Repairers, Computer Numerically Controlled Tool Operators, Construction Laborers, Conveyor Operators and Tenders, Cooks, Fast Food, Cooks, Institution and Cafeteria, Cooks, Short Order, Cutters and Trimmers, Hand, Cutting and Slicing Machine Setters, Operators, and Tenders, Cutting, Punching, and Press Machine Setters, Operators, and Tenders, Metal and Plastic, Dishwashers, Drywall and Ceiling Tile Installers, Electric Motor, Power Tool, and Related Repairers, Electrical and Electronic Equipment Assemblers, Electrical and Electronics Installers and Repairers, Transportation Equipment, Electrical and Electronics Repairers, Powerhouse, Substation, and Relay, Engine and Other Machine Assemblers, Etchers and Engravers, Extruding and Drawing Machine Setters, Operators, and Tenders, Metal and Plastic, Fallers, Farmworkers and Laborers, Crop, Nursery, and Greenhouse, Fence Erectors, Floor Layers, Except Carpet, Wood, and Hard Tiles, Food and Tobacco Roasting, Baking, and Drying Machine Operators and Tenders, Food Batchmakers, Food Cooking Machine Operators and Tenders, Food Preparation Workers, Foundry Mold and Coremakers, Furnace, Kiln, Oven, Drier, and Kettle Operators and Tenders, Gas Compressor and Gas Pumping Station Operators, Gas Plant Operators, Geothermal Technicians, Glaziers, Graders and Sorters, Agricultural Products, Grinding and Polishing Workers, Hand, Hazardous Materials Removal Workers, Heat Treating Equipment Setters, Operators, and Tenders, Metal and Plastic, Heavy and Tractor-Trailer Truck Drivers, Helpers--Carpenters, Helpers--Electricians, Helpers--Extraction Workers, Helpers--Painters, Paperhangers, Plasterers, and Stucco Masons, Helpers--Production Workers, Helpers--Roofers, Highway Maintenance Workers, Industrial Truck and Tractor Operators, Janitors and Cleaners, Except Maids and Housekeeping Cleaners, Landscaping and Groundskeeping Workers, Laundry and Dry-Cleaning Workers, Light Truck Drivers, Loading and Moving Machine Operators, Underground Mining, Locksmiths and Safe Repairers, Locomotive Engineers, Log Graders and Scalers, Machine Feeders and Offbearers, Maids and Housekeeping Cleaners, Manufactured Building and Mobile Home Installers, Milling and Planing Machine Setters, Operators, and Tenders, Metal and Plastic, Mobile Heavy Equipment Mechanics, Except Engines, Molders, Shapers, and Casters, Except Metal and Plastic, Motion Picture Projectionists, Office Machine Operators, Except Computer, Ophthalmic Laboratory Technicians, Outdoor Power Equipment and Other Small Engine Mechanics, Packaging and Filling Machine Operators and Tenders, Packers and Packagers, Hand, Painters, Construction and Maintenance, Paper Goods Machine Setters, Operators, and Tenders, Paperhangers, Parking Attendants, Parking Enforcement Workers, Paving, Surfacing, and Tamping Equipment Operators, Pest Control Workers, Pesticide Handlers, Sprayers, and Applicators, Vegetation, Petroleum Pump System Operators, Refinery Operators, and Gaugers, Plating Machine Setters, Operators, and Tenders, Metal and Plastic, Pourers and Casters, Metal, Power Plant Operators, Prepress Technicians and Workers, Pressers, Textile, Garment, and Related Materials, Print Binding and Finishing Workers, Printing Press Operators, Radio, Cellular, and Tower Equipment Installers and Repairers, Recycling and Reclamation Workers, Refuse and Recyclable Material Collectors, Rock Splitters, Quarry, Rolling Machine Setters, Operators, and Tenders, Metal and Plastic, Roof Bolters, Mining, Roofers, Sailors and Marine Oilers, Security and Fire Alarm Systems Installers, Segmental Pavers, Sewing Machine Operators, Shoe Machine Operators and Tenders, Signal and Track Switch Repairers, Slaughterers and Meat Packers, Solar Photovoltaic Installers, Solar Thermal Installers and Technicians, Structural Metal Fabricators and Fitters, Subway and Streetcar Operators, Tank Car, Truck, and Ship Loaders, Textile Cutting Machine Setters, Operators, and Tenders, Textile Knitting and Weaving Machine Setters, Operators, and Tenders, Textile Winding, Twisting, and Drawing Out Machine Setters, Operators, and Tenders, Timing Device Assemblers and Adjusters, Tire Builders, Tire Repairers and Changers, Tool Grinders, Filers, and Sharpeners, Water and Wastewater Treatment Plant and System Operators, Weatherization Installers and Technicians, Welding, Soldering, and Brazing Machine Setters, Operators, and Tenders, Wind Turbine Service Technicians,", "RCA": "Architectural and Civil Drafters, Glass Blowers, Molders, Benders, and Finishers, Shoe and Leather Workers and Repairers, Tile and Stone Setters,", "RCE": "Automotive Glass Installers and Repairers, Bailiffs, Bridge and Lock Tenders, Butchers and Meat Cutters, Couriers and Messengers, Forest Fire Inspectors and Prevention Specialists, Gambling Surveillance Officers and Gambling Investigators, Motorboat Operators, Nuclear Power Reactor Operators, Railroad Brake, Signal, and Switch Operators and Locomotive Firers, Security Guards, Ship Engineers, Team Assemblers, Transportation Inspectors,", "RCI": "Agricultural Inspectors, Aircraft Mechanics and Service Technicians, Airline Pilots, Copilots, and Flight Engineers, Aviation Inspectors, Bicycle Repairers, Biofuels Processing Technicians, Biomass Plant Technicians, Brickmasons and Blockmasons, Broadcast Technicians, Camera and Photographic Equipment Repairers, Carpenters, Civil Engineering Technologists and Technicians, Coating, Painting, and Spraying Machine Setters, Operators, and Tenders, Coil Winders, Tapers, and Finishers, Computer, Automated Teller, and Office Machine Repairers, Construction and Building Inspectors, Continuous Mining Machine Operators, Control and Valve Installers and Repairers, Except Mechanical Door, Cooling and Freezing Equipment Operators and Tenders, Crane and Tower Operators, Crushing, Grinding, and Polishing Machine Setters, Operators, and Tenders, Derrick Operators, Oil and Gas, Drilling and Boring Machine Tool Setters, Operators, and Tenders, Metal and Plastic, Electrical and Electronics Drafters, Electromechanical Equipment Assemblers, Embalmers, Extruding, Forming, Pressing, and Compacting Machine Setters, Operators, and Tenders, Farm Equipment Mechanics and Service Technicians, Fiberglass Laminators and Fabricators, Fire Inspectors and Investigators, Forest and Conservation Workers, Forging Machine Setters, Operators, and Tenders, Metal and Plastic, Heating, Air Conditioning, and Refrigeration Mechanics and Installers, Helpers--Installation, Maintenance, and Repair Workers, Home Appliance Repairers, Hydroelectric Plant Technicians, Insulation Workers, Mechanical, Layout Workers, Metal and Plastic, Machinists, Maintenance and Repair Workers, General, Maintenance Workers, Machinery, Mechanical Drafters, Medical Equipment Preparers, Millwrights, Mixing and Blending Machine Setters, Operators, and Tenders, Molding, Coremaking, and Casting Machine Setters, Operators, and Tenders, Metal and Plastic, Motorboat Mechanics and Service Technicians, Nanotechnology Engineering Technologists and Technicians, Nuclear Monitoring Technicians, Nuclear Technicians, Operating Engineers and Other Construction Equipment Operators, Patternmakers, Metal and Plastic, Pile Driver Operators, Plumbers, Pipefitters, and Steamfitters, Pump Operators, Except Wellhead Pumpers, Rail Car Repairers, Rail Yard Engineers, Dinkey Operators, and Hostlers, Refractory Materials Repairers, Except Brickmasons, Reinforcing Iron and Rebar Workers, Riggers, Sawing Machine Setters, Operators, and Tenders, Wood, Semiconductor Processing Technicians, Separating, Filtering, Clarifying, Precipitating, and Still Machine Setters, Operators, and Tenders, Service Unit Operators, Oil and Gas, Surveyors, Transportation Vehicle, Equipment and Systems Inspectors, Except Aviation, Watch and Clock Repairers, Welders, Cutters, Solderers, and Brazers, Wellhead Pumpers, Woodworking Machine Setters, Operators, and Tenders, Except Sawing,", "RCS": "Dining Room and Cafeteria Attendants and Bartender Helpers, Fast Food and Counter Workers, Magnetic Resonance Imaging Technologists,", "RE": "Athletes and Sports Competitors, Carpet Installers, Cement Masons and Concrete Finishers, Cooks, Restaurant, Driver/Sales Workers, Telecommunications Line Installers and Repairers,", "REC": "Barbers, Captains, Mates, and Pilots of Water Vessels, Computer Network Support Specialists, Correctional Officers and Jailers, Hoist and Winch Operators, Rotary Drill Operators, Oil and Gas, Shampooers, Telecommunications Engineering Specialists, Transit and Railroad Police, Transportation Security Screeners, Umpires, Referees, and Other Sports Officials,", "RES": "Manicurists and Pedicurists,", "RI": "Animal Breeders, Automotive Engineers, Fish and Game Wardens, Fuel Cell Engineers, Gem and Diamond Workers, Manufacturing Engineers, Mechanical Engineering Technologists and Technicians, Neurodiagnostic Technologists, Remote Sensing Scientists and Technologists, Robotics Technicians, Solar Energy Systems Engineers, Transportation Engineers, Veterinary Technologists and Technicians,", "RIA": "Patternmakers, Wood,", "RIC": "Aerospace Engineering and Operations Technologists and Technicians, Agricultural Technicians, Audio and Video Technicians, Automotive Engineering Technicians, Automotive Service Technicians and Mechanics, Avionics Technicians, Biological Technicians, Cartographers and Photogrammetrists, Civil Engineers, Computer User Support Specialists, Dental Laboratory Technicians, Earth Drillers, Except Oil and Gas, Electrical and Electronic Engineering Technologists and Technicians, Electrical and Electronics Repairers, Commercial and Industrial Equipment, Electrical Power-Line Installers and Repairers, Electricians, Electro-Mechanical and Mechatronics Technologists and Technicians, Electronic Equipment Installers and Repairers, Motor Vehicles, Elevator and Escalator Installers and Repairers, Endoscopy Technicians, Environmental Engineering Technologists and Technicians, Explosives Workers, Ordnance Handling Experts, and Blasters, Food Science Technicians, Geological Technicians, Except Hydrologic Technicians, Industrial Machinery Mechanics, Lathe and Turning Machine Tool Setters, Operators, and Tenders, Metal and Plastic, Logging Equipment Operators, Medical and Clinical Laboratory Technicians, Medical Equipment Repairers, Metal-Refining Furnace Operators and Tenders, Model Makers, Metal and Plastic, Non-Destructive Testing Specialists, Photonics Technicians, Power Distributors and Dispatchers, Precision Agriculture Technicians, Radio Frequency Identification Device Specialists, Recreational Vehicle Service Technicians, Remote Sensing Technicians, Stationary Engineers and Boiler Operators, Structural Iron and Steel Workers, Telecommunications Equipment Installers and Repairers, Except Line Installers, Tool and Die Makers,", "RIE": "Commercial Pilots, Conservation Scientists, Fishing and Hunting Workers, Forest and Conservation Technicians, Foresters, Range Managers, Traffic Technicians, Wind Energy Engineers,", "RIS": "Cardiovascular Technologists and Technicians, Medical Appliance Technicians,", "RS": "Ambulance Drivers and Attendants, Except Emergency Medical Technicians, Bus Drivers, Transit and Intercity, Lifeguards, Ski Patrol, and Other Recreational Protective Service Workers,", "RSC": "Animal Control Workers, Radiologic Technologists and Technicians, Surgical Technologists,", "RSE": "Firefighters, Passenger Attendants,", "RSI": "Anesthesiologist Assistants, Oral and Maxillofacial Surgeons, Veterinary Assistants and Laboratory Animal Caretakers,", "S": "Career/Technical Education Teachers, Secondary School, Educational, Guidance, and Career Counselors and Advisors, Recreation and Fitness Studies Teachers, Postsecondary,", "SA": "Architecture Teachers, Postsecondary, Art, Drama, and Music Teachers, Postsecondary, Childcare Workers, Communications Teachers, Postsecondary, Kindergarten Teachers, Except Special Education, Middle School Teachers, Except Special and Career/Technical Education, Preschool Teachers, Except Special Education, Recreational Therapists, Special Education Teachers, Middle School, Special Education Teachers, Preschool,", "SAC": "Career/Technical Education Teachers, Middle School, Elementary School Teachers, Except Special Education, Training and Development Specialists,", "SAE": "Adult Basic Education, Adult Secondary Education, and English as a Second Language Instructors, Nannies, Secondary School Teachers, Except Special and Career/Technical Education, Self-Enrichment Teachers,", "SAI": "Art Therapists, Education Teachers, Postsecondary, English Language and Literature Teachers, Postsecondary, Foreign Language and Literature Teachers, Postsecondary, Marriage and Family Therapists, Music Therapists, Philosophy and Religion Teachers, Postsecondary, Substance Abuse and Behavioral Disorder Counselors,", "SC": "Speech-Language Pathology Assistants, Teaching Assistants, Postsecondary,", "SCE": "Eligibility Interviewers, Government Programs, Ushers, Lobby Attendants, and Ticket Takers,", "SCR": "Medical Assistants, Nursing Assistants,", "SE": "Arbitrators, Mediators, and Conciliators, Child, Family, and School Social Workers, Concierges, Emergency Management Directors, Health Education Specialists, Patient Representatives, Tour Guides and Escorts,", "SEA": "Clergy, Political Science Teachers, Postsecondary, Recreation Workers,", "SEC": "Education and Childcare Administrators, Preschool and Daycare, Equal Opportunity Representatives and Officers, Locker Room, Coatroom, and Dressing Room Attendants, Probation Officers and Correctional Treatment Specialists, Residential Advisors, Waiters and Waitresses,", "SEI": "Business Teachers, Postsecondary,", "SER": "Crossing Guards and Flaggers, Funeral Attendants, Psychiatric Technicians,", "SI": "Advanced Practice Psychiatric Nurses, Anthropology and Archeology Teachers, Postsecondary, Atmospheric, Earth, Marine, and Space Sciences Teachers, Postsecondary, Biological Science Teachers, Postsecondary, Criminal Justice and Law Enforcement Teachers, Postsecondary, Economics Teachers, Postsecondary, Geography Teachers, Postsecondary, Health Informatics Specialists, Health Specialties Teachers, Postsecondary, Healthcare Social Workers, Hospitalists, Nurse Midwives, Nursing Instructors and Teachers, Postsecondary, Occupational Therapists, Physics Teachers, Postsecondary, Rehabilitation Counselors, Social Work Teachers, Postsecondary, Special Education Teachers, Secondary School, Tutors,", "SIA": "Area, Ethnic, and Cultural Studies Teachers, Postsecondary, Environmental Science Teachers, Postsecondary, Family and Consumer Sciences Teachers, Postsecondary, Genetic Counselors, History Teachers, Postsecondary, Mathematical Science Teachers, Postsecondary, Mental Health and Substance Abuse Social Workers, Mental Health Counselors, Psychology Teachers, Postsecondary, Sociology Teachers, Postsecondary, Speech-Language Pathologists,", "SIC": "Computer Science Teachers, Postsecondary, Library Science Teachers, Postsecondary, Registered Nurses,", "SIE": "Community Health Workers, Instructional Coordinators, Law Teachers, Postsecondary,", "SIR": "Acute Care Nurses, Agricultural Sciences Teachers, Postsecondary, Chemistry Teachers, Postsecondary, Chiropractors, Critical Care Nurses, Dietetic Technicians, Engineering Teachers, Postsecondary, Exercise Physiologists, Forestry and Conservation Science Teachers, Postsecondary, Hearing Aid Specialists, Low Vision Therapists, Orientation and Mobility Specialists, and Vision Rehabilitation Therapists, Nurse Practitioners, Physical Therapists, Physician Assistants, Preventive Medicine Physicians, Respiratory Therapists,", "SR": "Adapted Physical Education Specialists, Career/Technical Education Teachers, Postsecondary, Home Health Aides, Licensed Practical and Licensed Vocational Nurses, Massage Therapists, Midwives, Occupational Therapy Aides, Occupational Therapy Assistants, Physical Therapist Aides,", "SRA": "Park Naturalists,", "SRC": "Dental Hygienists, Personal Care Aides, Psychiatric Aides, Radiation Therapists,", "SRE": "Coaches and Scouts, Exercise Trainers and Group Fitness Instructors, Farm and Home Management Educators, Food Servers, Nonrestaurant,", "SRI": "Acupuncturists, Athletic Trainers, Orthotists and Prosthetists, Physical Therapist Assistants,"}};

// ============================================================
// 108 SOAL STANDAR HOLLAND (Sesuai Sheet Soal Excel)
// ============================================================
var SOAL_STANDAR_HOLLAND = [{"no": 1, "code": "A8", "teks": "Senang menonton drama", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 2, "code": "S5", "teks": "Senang Melatih Orang", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 3, "code": "A11", "teks": "Mampu bermain dalam drama / berakting", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 4, "code": "I5", "teks": "Suka membaca mengenai topik-topik khusus atas keinginan sendiri", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 5, "code": "I7", "teks": "Mampu melakukan percobaan atau penelitian ilmiah", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 6, "code": "A15", "teks": "Artistik", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 7, "code": "I8", "teks": "Mampu memprogram komputer untuk mempelajari masalah ilmiah", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 8, "code": "R15", "teks": "Tertarik menjadi pengawas konstruksi bangunan", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 9, "code": "I15", "teks": "Suka dalam memecahkan soal eksak", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 10, "code": "C11", "teks": "Mampu melakukan dan menyukai tugas administratif", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 11, "code": "E1", "teks": "Bisa mempengaruhi dan membujuk orang lain", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 12, "code": "C9", "teks": "Prosedural (Mengikuti aturan)", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 13, "code": "C10", "teks": "Suka menggunakan aplikasi pencatatan keuangan", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 14, "code": "R2", "teks": "Suka memperbaiki motor", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 15, "code": "A18", "teks": "Bisa mengekspresikan diri secara kreatif", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 16, "code": "C17", "teks": "Ingin bekerja dibagian Staff keuangan", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 17, "code": "I11", "teks": "Mampu menyebutkan makanan yang memiliki protein tinggi", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 18, "code": "S12", "teks": "Suka melakukan pekerjaan social", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 19, "code": "E7", "teks": "Mampu mengetahui bagaimana menjadi pemimpin yang baik/ berhasil", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 20, "code": "R14", "teks": "Tidak suka menyatakan perasaan", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 21, "code": "E17", "teks": "Percaya diri", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 22, "code": "I6", "teks": "Suka menerapkan matematika dalam masalah praktis", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 23, "code": "A4", "teks": "Emosional", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 24, "code": "R3", "teks": "Menyukai memperbaiki peralatan mekanik", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 25, "code": "S13", "teks": "Mampu menghibur dan menemani orang yang lebih tua dari saya", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 26, "code": "E11", "teks": "Senang memulai proyek baru", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 27, "code": "I10", "teks": "Suka kegiatan akademis", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 28, "code": "S2", "teks": "Ramah", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 29, "code": "C16", "teks": "Efisien dan Terstruktur", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 30, "code": "R4", "teks": "Suka menggunakan perkakas bengkel dan mesin", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 31, "code": "S4", "teks": "Suka menjaga / mengurus mengawasi anak-anak", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 32, "code": "I2", "teks": "Senang menganalisa dan mengevaluasi", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 33, "code": "R7", "teks": "Senang melatih binatang", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 34, "code": "A17", "teks": "Tertarik menjadi aktor/ aktris", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 35, "code": "A10", "teks": "Mampu memainkan alat musik", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 36, "code": "S16", "teks": "Tertarik menjadi konselor masalah pribadi", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 37, "code": "A2", "teks": "Ingin menjadi pemain dalam kelompok musik", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 38, "code": "I12", "teks": "Tertarik menjadi ahli biologi/ hayati", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 39, "code": "S17", "teks": "Tertarik menjadi pekerja sosial", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 40, "code": "E3", "teks": "Energetik", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 41, "code": "S8", "teks": "Suka bekerja dalam kelompok", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 42, "code": "I1", "teks": "Suka membaca buku/ majalah ilmiah", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 43, "code": "R11", "teks": "Suka mengutak-atik mesin/alat elektronik", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 44, "code": "R12", "teks": "Suka kegiatan di luar ruangan", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 45, "code": "A16", "teks": "Senang memotret atau videographer", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 46, "code": "A7", "teks": "Suka membuat lukisan/ foto orang", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 47, "code": "A13", "teks": "Senang bekerja di situasi yang bebas", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 48, "code": "C12", "teks": "Membuat Hari Saya Terstruktur", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 49, "code": "S14", "teks": "Mampu menarik perhatian orang untuk menceritakan masalah mereka", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 50, "code": "C14", "teks": "Memperhatikan Detailnya", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 51, "code": "C15", "teks": "Sering diminta mengumpulkan data", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 52, "code": "A1", "teks": "Suka membuat sketsa, Mendesign atau melukis", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 53, "code": "S11", "teks": "Mampu mengajar orang dewasa dengan mudah", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 54, "code": "S9", "teks": "Antusias berparisipasi dalam pencarian dana/ amal", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 55, "code": "S6", "teks": "Mampu mudah berbicara dengan semua orang", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 56, "code": "I4", "teks": "Suka mempelajari teori ilmiah", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 57, "code": "I18", "teks": "Tertarik menjadi pekerja riset ilmiah", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 58, "code": "E13", "teks": "Mampu mengatur pekerjaan orang lain", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 59, "code": "R10", "teks": "Bisa menguasai diri", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 60, "code": "E10", "teks": "Mampu membuat kelompok sosial/ kerja berjalan dengan baik", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 61, "code": "R17", "teks": "Tertarik menjadi insinyur otomotif", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 62, "code": "C3", "teks": "Suka membuat catatan pengeluaran yang terperinci", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 63, "code": "I9", "teks": "Kompleks", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 64, "code": "E9", "teks": "Mampu mengelola usaha kecil", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 65, "code": "A6", "teks": "Rumit", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 66, "code": "R13", "teks": "Tertarik menjadi spesialis perikanan/ margasatwa", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 67, "code": "E5", "teks": "Suka memimpin kelompok dalam meraih tujuan tertentu", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 68, "code": "E2", "teks": "Suka menjual suatu barang", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 69, "code": "E16", "teks": "Tegas", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 70, "code": "C13", "teks": "Mampu mencatat dengan cermat pembayaran/ penjualan", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 71, "code": "I3", "teks": "Suka mengerjakan proyek ilmiah", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 72, "code": "S1", "teks": "Bisa menyatakan perasaan dengan jelas", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 73, "code": "R6", "teks": "Menyusun Sesuatu Atau Merakit Model", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 74, "code": "I17", "teks": "Senang mengamati", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 75, "code": "E15", "teks": "Mampu memengaruhi orang lain supaya melakukan sesuatu dengan caranya", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 76, "code": "A14", "teks": "Senang mengerjakan kerajinan tangan", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 77, "code": "R9", "teks": "Suka bekerja menggunakan tangan", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 78, "code": "S3", "teks": "Suka membantu orang lain dengan masalah pribadinya", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 79, "code": "R16", "teks": "Senang kegiatan fisik", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 80, "code": "I14", "teks": "Tertarik menjadi ahli teknisi laboratorium medis", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 81, "code": "A5", "teks": "Ingin menjadi seniman", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 82, "code": "R8", "teks": "Mampu melakukan perbaikan kecil pada pipa air, kran, dan lain-lain", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 83, "code": "R5", "teks": "Mampu membuat gambar dengan skala", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 84, "code": "C6", "teks": "Prosedural (Suka mengikuti aturan)", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 85, "code": "C8", "teks": "Suka mengarsip surat dan berkas-berkas lain", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 86, "code": "E8", "teks": "Mampu menjadi seorang pembicara di depan umum yang baik", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 87, "code": "S7", "teks": "Mampu memimpin diskusi kelompok", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 88, "code": "C1", "teks": "Suka melakukan pekerjaan surat menyurat/ masalah perkantoran", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 89, "code": "E6", "teks": "Senang bertemu orang baru", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 90, "code": "C7", "teks": "Bisa bekerja dengan baik dalam system /aturan", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 91, "code": "A3", "teks": "Suka merancang perabotan, pakaian atau poster", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 92, "code": "E18", "teks": "Tertarik menjadi pembawa acara / MC", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 93, "code": "I16", "teks": "Tertarik menjadi ilmuwan peneliti", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 94, "code": "R1", "teks": "Suka memperbaiki alat-alat listrik", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 95, "code": "E12", "teks": "Terbuka", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 96, "code": "C18", "teks": "Tertarik menjadi penaksir biaya", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 97, "code": "S15", "teks": "Suka mendamaikan orang", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 98, "code": "I13", "teks": "Suka mengerjakan teka-teki", "kode": "I", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 99, "code": "S10", "teks": "Mampu mengajar anak-anak dengan mudah", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 100, "code": "C5", "teks": "Suka membuat daftar inventaris dari persediaan/ produk", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 101, "code": "R18", "teks": "Suka olahraga", "kode": "R", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 102, "code": "A9", "teks": "Bisa menulis karya sastra", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 103, "code": "S18", "teks": "Tertarik menjadi konselor kejuruan dan pekerjaan", "kode": "S", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 104, "code": "A12", "teks": "Sensitif", "kode": "A", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 105, "code": "E4", "teks": "Senang petualangan / mengambil resiko", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 106, "code": "E14", "teks": "Ambisius dan cenderung berbicara apa adanya (Spontan)", "kode": "E", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 107, "code": "C2", "teks": "Penurut", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}, {"no": 108, "code": "C4", "teks": "Suka menyusun sistem pengarsipan", "kode": "C", "jenjang": "SEMUA", "opsiA": "Ya", "opsiB": "Tidak", "nilaiA": 1, "nilaiB": 0}];

// ============================================================
// ENTRY POINT - doGet
// ============================================================
function doGet(e) {
  e = e || { parameter: {} };
  var page = (e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : 'login';

  var pageMap = {
    'login'        : 'Login',
    'admin-login'  : 'LoginAdmin',
    'biodata'      : 'Biodata',
    'instrumen'    : 'Instrumen',
    'hasil'        : 'Hasil',
    'admin'        : 'Admin',
    'laporan'      : 'Laporan',
    'laporan-siswa': 'LaporanSiswa'
  };

  var viewName = pageMap[page] || 'Login';
  var template = HtmlService.createTemplateFromFile(viewName);
  template.page = page;

  return template.evaluate()
    .setTitle('FOMO - Report Holland Test (RIASEC)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// UTILITY: Include file HTML
// ============================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// UTILITY: Format Tanggal Aman (Anti-Crash)
// ============================================================
function safeFormatDate(val, formatStr) {
  if (!val) return '';
  formatStr = formatStr || 'dd/MM/yyyy HH:mm';
  try {
    var d;
    if (val instanceof Date) {
      d = val;
    } else {
      d = new Date(val);
    }
    if (isNaN(d.getTime())) {
      return String(val);
    }
    return Utilities.formatDate(d, 'Asia/Jakarta', formatStr);
  } catch(e) {
    return String(val || '');
  }
}

// ============================================================
// UTILITY: Ambil/Buat Spreadsheet OTOMATIS
// ============================================================
function getSpreadsheetId() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');

  if (id) {
    try {
      SpreadsheetApp.openById(id);
      return id;
    } catch(e) {
      props.deleteProperty('SPREADSHEET_ID');
    }
  }

  var ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  var newId = ss.getId();
  props.setProperty('SPREADSHEET_ID', newId);

  try {
    var file = DriveApp.getFileById(newId);
    var rootId = DriveApp.getRootFolder().getId();
    DriveApp.getFolderById(rootId).addFile(file);
  } catch(e) { }

  return newId;
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
}

function getSheet(name) {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    initSheetHeader(sh, name);
  }
  return sh;
}

function getSpreadsheetUrl() {
  try {
    return { success: true, url: 'https://docs.google.com/spreadsheets/d/' + getSpreadsheetId() };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Map header kolom ke index (case-insensitive)
function getColumnMap(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var c = 0; c < headers.length; c++) {
    var key = headers[c] ? headers[c].toString().trim().toLowerCase().replace(/[^a-z0-9_]/g, '_') : '';
    if (key) map[key] = c;
  }
  return map;
}

// Inisialisasi header tiap sheet jika baru
function initSheetHeader(sh, name) {
  var headers = {
    'Token'   : ['Token', 'Nama_Wali', 'No_HP_Wali', 'Jenjang', 'Asal_Sekolah', 'Status', 'Tgl_Dibuat', 'Tgl_Dipakai'],
    'Biodata' : ['Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'No_HP', 'Timestamp'],
    'Soal'    : ['ID_Soal', 'Teks_Soal', 'Kode_RIASEC', 'Jenjang', 'Opsi_A', 'Opsi_B', 'Nilai_A', 'Nilai_B', 'Aktif'],
    'Jawaban' : ['Token', 'ID_Soal', 'Jawaban', 'Nilai', 'Timestamp'],
    'Hasil'   : ['Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'Skor_R', 'Skor_I', 'Skor_A', 'Skor_S', 'Skor_E', 'Skor_C', 'Bidang_Utama', 'Bidang_2', 'Bidang_3', 'Kode_Holland', 'Konsistensi', 'Timestamp'],
    'Admin'   : ['Username', 'Password', 'Role', 'Tgl_Dibuat']
  };

  if (headers[name] && sh.getLastRow() === 0) {
    sh.appendRow(headers[name]);
    sh.getRange(1, 1, 1, headers[name].length)
      .setBackground('#1565C0')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

// ============================================================
// FUNGSI INTERPRETASI HOLLAND (Persis Rumus Excel Hasil & Referensi)
// ============================================================
function getInterpretasiHolland(top3Kode) {
  top3Kode = (top3Kode || 'RIA').toString().trim().toUpperCase();
  var c1 = top3Kode.charAt(0) || 'R';
  var c2 = top3Kode.charAt(1) || 'I';
  var c3 = top3Kode.charAt(2) || 'A';
  var code2 = c1 + c2;
  var code3 = c1 + c2 + c3;

  var konsistensi = (REFERENSI_RIASEC.konsistensi && REFERENSI_RIASEC.konsistensi[code2]) ?
                    REFERENSI_RIASEC.konsistensi[code2] : 'Sedang';

  var combo2 = REFERENSI_RIASEC.combo2 && REFERENSI_RIASEC.combo2[code2];
  var single1 = (REFERENSI_RIASEC.single && REFERENSI_RIASEC.single[c1]) || { deskripsi: '', profesi: '' };

  var deskripsi = '';
  if (combo2 && combo2.deskripsi) {
    deskripsi = combo2.deskripsi;
  } else {
    deskripsi = single1.deskripsi || '';
  }

  var profesiTerkait = '';
  if (combo2 && combo2.profesi) {
    profesiTerkait = combo2.profesi;
  }
  if (single1.profesi) {
    profesiTerkait = (profesiTerkait ? profesiTerkait + ' ' : '') + single1.profesi;
  }

  var multi = REFERENSI_RIASEC.profesi_kombinasi || {};
  var profesi3 = multi[code3] || '';
  var profesi2 = multi[code2] || '';

  return {
    top3Kode: code3,
    code2: code2,
    c1: c1,
    c2: c2,
    c3: c3,
    nama1: BIDANG_MAP[c1] || c1,
    nama2: BIDANG_MAP[c2] || c2,
    nama3: BIDANG_MAP[c3] || c3,
    konsistensi: konsistensi,
    deskripsi: deskripsi,
    profesiTerkait: profesiTerkait,
    profesiKombinasi3: profesi3,
    profesiKombinasi2: profesi2
  };
}

// ============================================================
// SETUP AWAL
// ============================================================
function setupAwal() {
  var ss = getSpreadsheet();
  var sheets = [SHEET_TOKEN, SHEET_BIODATA, SHEET_SOAL, SHEET_JAWABAN, SHEET_HASIL, SHEET_ADMIN];

  sheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      initSheetHeader(sh, name);
    }
  });

  // Admin default
  var adminSheet = getSheet(SHEET_ADMIN);
  if (adminSheet.getLastRow() <= 1) {
    adminSheet.appendRow(['admin', 'fomo2024', 'superadmin', new Date()]);
  }

  // Token contoh
  var tokenSheet = getSheet(SHEET_TOKEN);
  if (tokenSheet.getLastRow() <= 1) {
    tokenSheet.appendRow(['SD-SAMPLE', 'Admin Guru', '081234567890', 'SD', 'SDN Contoh 01', 'aktif', new Date(), '']);
    tokenSheet.appendRow(['TK-SAMPLE', 'Admin Guru', '081234567890', 'TK', 'TK Contoh 01', 'aktif', new Date(), '']);
    tokenSheet.appendRow(['UM-SAMPLE', 'Umum', '081234567890', 'SEMUA', 'Umum', 'aktif', new Date(), '']);
  }

  // Isi 108 soal standar Excel jika kosong
  var soalSheet = getSheet(SHEET_SOAL);
  if (soalSheet.getLastRow() <= 1) {
    insertSoalContoh(soalSheet);
  }

  return {
    success: true,
    message: 'Setup database berhasil dengan 108 soal standar Holland.',
    url: 'https://docs.google.com/spreadsheets/d/' + ss.getId()
  };
}

// ============================================================
// FUNGSI SINKRONISASI & PERBAIKAN DATABASE (Jalankan Sekali)
// ============================================================
function sinkronkanDatabase() {
  var ss = getSpreadsheet();
  
  // 1. Standarisasi Header Baku Semua Sheet
  var headers = {
    'Token'   : ['Token', 'Nama_Wali', 'No_HP_Wali', 'Jenjang', 'Asal_Sekolah', 'Status', 'Tgl_Dibuat', 'Tgl_Dipakai'],
    'Biodata' : ['Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'No_HP', 'Timestamp'],
    'Soal'    : ['ID_Soal', 'Teks_Soal', 'Kode_RIASEC', 'Jenjang', 'Opsi_A', 'Opsi_B', 'Nilai_A', 'Nilai_B', 'Aktif'],
    'Jawaban' : ['Token', 'ID_Soal', 'Jawaban', 'Nilai', 'Timestamp'],
    'Hasil'   : ['Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'Skor_R', 'Skor_I', 'Skor_A', 'Skor_S', 'Skor_E', 'Skor_C', 'Bidang_Utama', 'Bidang_2', 'Bidang_3', 'Kode_Holland', 'Konsistensi', 'Timestamp'],
    'Admin'   : ['Username', 'Password', 'Role', 'Tgl_Dibuat']
  };

  Object.keys(headers).forEach(function(shName) {
    var sh = ss.getSheetByName(shName);
    if (!sh) {
      sh = ss.insertSheet(shName);
    }
    var hList = headers[shName];
    sh.getRange(1, 1, 1, hList.length).setValues([hList]);
    sh.getRange(1, 1, 1, hList.length)
      .setBackground('#1565C0')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  });

  // 2. Rapikan sheet Hasil jika ada baris lama yang kolomnya belum lengkap
  var hasilSh = ss.getSheetByName('Hasil');
  if (hasilSh && hasilSh.getLastRow() > 1) {
    var lastRow = hasilSh.getLastRow();
    var lastCol = Math.max(hasilSh.getLastColumn(), 18);
    var data = hasilSh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var token = row[0] ? row[0].toString().trim().toUpperCase() : '';
      if (!token) continue;

      var kodeHolland = row[15];
      if (!kodeHolland || kodeHolland.toString().trim() === '') {
        var skor = {
          R: parseFloat(row[6]) || 0,
          I: parseFloat(row[7]) || 0,
          A: parseFloat(row[8]) || 0,
          S: parseFloat(row[9]) || 0,
          E: parseFloat(row[10]) || 0,
          C: parseFloat(row[11]) || 0
        };
        var ranked = ['R', 'I', 'A', 'S', 'E', 'C'].map(function(k) {
          return { kode: k, skor: skor[k], bidang: BIDANG_MAP[k] || k };
        }).sort(function(a, b) { return b.skor - a.skor; });

        var top3 = ranked.slice(0, 3).map(function(r) { return r.kode; }).join('');
        var interp = getInterpretasiHolland(top3);

        hasilSh.getRange(i + 2, 13).setValue(ranked[0].bidang);
        hasilSh.getRange(i + 2, 14).setValue(ranked[1] ? ranked[1].bidang : '');
        hasilSh.getRange(i + 2, 15).setValue(ranked[2] ? ranked[2].bidang : '');
        hasilSh.getRange(i + 2, 16).setValue(top3);
        hasilSh.getRange(i + 2, 17).setValue(interp.konsistensi);
      }
    }
  }

  return 'Database berhasil disinkronkan dan header diperbarui.';
}


function insertSoalContoh(sheet) {
  var rows = SOAL_STANDAR_HOLLAND.map(function(s) {
    var idStr = 'S' + ('000' + s.no).slice(-3);
    return [
      idStr,
      s.teks,
      s.kode,
      s.jenjang || 'SEMUA',
      s.opsiA || 'Ya',
      s.opsiB || 'Tidak',
      s.nilaiA !== undefined ? s.nilaiA : 1,
      s.nilaiB !== undefined ? s.nilaiB : 0,
      'aktif'
    ];
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// ============================================================
// LOGIN TOKEN
// ============================================================
function loginToken(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var storedToken = row[0].toString().trim().toUpperCase();
      var status = (row[5] || '').toString().trim().toLowerCase();

      if (storedToken === token) {
        if (status === 'nonaktif') {
          return { success: false, message: 'Token sudah tidak aktif. Hubungi administrator.' };
        }

        // Ambil data hasil jika ada
        var hasilAda = cekHasilByToken(token);
        var bio = cekBiodata(token);

        if (hasilAda) {
          return {
            success    : true,
            status     : 'sudah_selesai',
            token      : token,
            jenjang    : row[3],
            sekolah    : row[4],
            namaPeserta: (bio.ada && bio.nama) ? bio.nama : (hasilAda.nama || ''),
            hasil      : hasilAda
          };
        }

        if (status === 'aktif') {
          sheet.getRange(i + 1, 6).setValue('dipakai');
          sheet.getRange(i + 1, 8).setValue(new Date());
        }

        return {
          success: true,
          status : bio.ada ? 'lanjut' : 'baru',
          token  : token,
          jenjang: row[3],
          sekolah: row[4],
          biodata: bio
        };
      }
    }
    return { success: false, message: 'Token "' + token + '" tidak ditemukan. Periksa kembali token Anda.' };
  } catch(err) {
    return { success: false, message: 'Terjadi kesalahan login: ' + err.message };
  }
}

// ============================================================
// RESET TOKEN UNTUK TES ULANG (Mulai dari 0)
// ============================================================
function resetToken(token) {
  try {
    token = token.toString().trim().toUpperCase();

    // 1. Update status di sheet Token
    var tokenSheet = getSheet(SHEET_TOKEN);
    var tData = tokenSheet.getDataRange().getValues();
    for (var i = 1; i < tData.length; i++) {
      if (tData[i][0].toString().trim().toUpperCase() === token) {
        tokenSheet.getRange(i + 1, 6).setValue('aktif');
        tokenSheet.getRange(i + 1, 8).setValue('');
        break;
      }
    }

    // 2. Hapus dari sheet Hasil
    var hasilSheet = getSheet(SHEET_HASIL);
    var hData = hasilSheet.getDataRange().getValues();
    for (var h = hData.length - 1; h >= 1; h--) {
      if (hData[h][0].toString().trim().toUpperCase() === token) {
        hasilSheet.deleteRow(h + 1);
      }
    }

    // 3. Hapus dari sheet Jawaban
    var jwbSheet = getSheet(SHEET_JAWABAN);
    var jData = jwbSheet.getDataRange().getValues();
    for (var j = jData.length - 1; j >= 1; j--) {
      if (jData[j][0].toString().trim().toUpperCase() === token) {
        jwbSheet.deleteRow(j + 1);
      }
    }

    // 4. Hapus dari sheet Biodata
    var bioSheet = getSheet(SHEET_BIODATA);
    var bData = bioSheet.getDataRange().getValues();
    for (var b = bData.length - 1; b >= 1; b--) {
      if (bData[b][0].toString().trim().toUpperCase() === token) {
        bioSheet.deleteRow(b + 1);
      }
    }

    return { success: true, message: 'Token ' + token + ' berhasil di-reset. Anda dapat mulai mengerjakan dari awal.' };
  } catch(err) {
    return { success: false, message: 'Gagal mereset token: ' + err.message };
  }
}

// ============================================================
// LOGIN ADMIN
// ============================================================
function loginAdmin(username, password) {
  try {
    var sheet = getSheet(SHEET_ADMIN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(username).trim() && String(data[i][1]).trim() === String(password).trim()) {
        return { success: true, role: data[i][2], username: username };
      }
    }
    return { success: false, message: 'Username atau password salah.' };
  } catch(err) {
    return { success: false, message: 'Terjadi kesalahan: ' + err.message };
  }
}

// ============================================================
// BIODATA SISWA
// ============================================================
function cekBiodata(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_BIODATA);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { ada: false };

    var col = getColumnMap(sheet);
    var cToken = col['token'] !== undefined ? col['token'] : 0;
    var cNama  = col['nama_siswa'] !== undefined ? col['nama_siswa'] : 1;
    var cJk    = col['jenis_kelamin'] !== undefined ? col['jenis_kelamin'] : -1;
    var cTgl   = col['tanggal_lahir'] !== undefined ? col['tanggal_lahir'] : -1;
    var cPend  = col['pendidikan'] !== undefined ? col['pendidikan'] : (col['jenjang'] !== undefined ? col['jenjang'] : 3);
    var cSek   = col['asal_sekolah'] !== undefined ? col['asal_sekolah'] : 4;
    var cHp    = col['no_hp'] !== undefined ? col['no_hp'] : 5;

    for (var i = 1; i < data.length; i++) {
      if (data[i][cToken] && data[i][cToken].toString().trim().toUpperCase() === token) {
        var row = data[i];
        return {
          ada         : true,
          token       : token,
          nama        : row[cNama] || '',
          jenisKelamin: cJk >= 0 ? (row[cJk] || 'Laki-laki') : (row[2] || 'Laki-laki'),
          tanggalLahir: cTgl >= 0 ? safeFormatDate(row[cTgl], 'yyyy-MM-dd') : '',
          tanggalLahirText: cTgl >= 0 ? safeFormatDate(row[cTgl], 'dd MMM yyyy') : '',
          pendidikan  : row[cPend] || 'SD',
          jenjang     : row[cPend] || 'SD',
          sekolah     : cSek >= 0 ? (row[cSek] || '') : '',
          noHp        : cHp >= 0 ? (row[cHp] || '') : ''
        };
      }
    }
    return { ada: false };
  } catch(err) {
    return { ada: false };
  }
}

function simpanBiodata(token, nama, jenisKelamin, tanggalLahir, pendidikan, sekolah, noHp) {
  try {
    token = token.toString().trim().toUpperCase();
    nama = String(nama || '').trim();
    jenisKelamin = String(jenisKelamin || 'Laki-laki').trim();
    tanggalLahir = tanggalLahir ? new Date(tanggalLahir) : '';
    pendidikan = String(pendidikan || 'SD').trim();
    sekolah = String(sekolah || '').trim();
    noHp = String(noHp || '').trim();

    var sheet = getSheet(SHEET_BIODATA);
    var data = sheet.getDataRange().getValues();
    var col = getColumnMap(sheet);
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        rowIndex = i + 1;
        break;
      }
    }

    // Jika header belum punya jenis kelamin, sesuaikan header
    if (col['jenis_kelamin'] === undefined && sheet.getLastRow() >= 1) {
      sheet.getRange(1, 1, 1, 8).setValues([[
        'Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'No_HP', 'Timestamp'
      ]]).setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold');
    }

    var rowValues = [token, nama, jenisKelamin, tanggalLahir, pendidikan, sekolah, noHp, new Date()];

    if (rowIndex > 1) {
      // Update data jika sudah ada
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
      return { success: true, message: 'Biodata berhasil diperbarui.', biodata: cekBiodata(token) };
    } else {
      sheet.appendRow(rowValues);
      return { success: true, message: 'Biodata berhasil disimpan.', biodata: cekBiodata(token) };
    }
  } catch(err) {
    return { success: false, message: 'Gagal menyimpan biodata: ' + err.message };
  }
}

// ============================================================
// SOAL
// ============================================================
function getSoal(jenjang) {
  try {
    var sheet = getSheet(SHEET_SOAL);
    var data = sheet.getDataRange().getValues();
    var soalList = [];

    // Jika soal masih kosong di sheet, isi otomatis dari 108 soal standar
    if (data.length <= 1) {
      insertSoalContoh(sheet);
      data = sheet.getDataRange().getValues();
    }

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var soalJenjang = (row[3] || '').toString().trim().toUpperCase();
      var aktif = (row[8] || '').toString().trim().toLowerCase();

      if (aktif !== 'tidak' && aktif !== 'false' && aktif !== '0') {
        if (!jenjang || jenjang === 'SEMUA' || soalJenjang === 'SEMUA' || soalJenjang === jenjang.toUpperCase()) {
          soalList.push({
            id       : row[0].toString(),
            no       : i,
            teks     : (row[1] || '').toString(),
            kode     : (row[2] || 'R').toString().trim().toUpperCase(),
            jenjang  : (row[3] || 'SEMUA').toString(),
            opsiA    : (row[4] || 'Ya').toString(),
            opsiB    : (row[5] || 'Tidak').toString(),
            nilaiA   : parseFloat(row[6]) !== undefined && !isNaN(parseFloat(row[6])) ? parseFloat(row[6]) : 1,
            nilaiB   : parseFloat(row[7]) !== undefined && !isNaN(parseFloat(row[7])) ? parseFloat(row[7]) : 0
          });
        }
      }
    }

    return { success: true, soal: soalList, total: soalList.length };
  } catch(err) {
    return { success: false, message: 'Gagal memuat soal: ' + err.message, soal: [] };
  }
}

// ============================================================
// SIMPAN JAWABAN & HITUNG SKOR RIASEC
// ============================================================
function simpanJawabanDanHitung(token, jawabanList) {
  try {
    token = token.toString().trim().toUpperCase();
    if (!jawabanList || !jawabanList.length) {
      return { success: false, message: 'Daftar jawaban kosong.' };
    }

    var ts = new Date();

    // 1. Simpan jawaban ke sheet Jawaban (bersihkan jawaban lama untuk token ini jika ada)
    var jawabanSheet = getSheet(SHEET_JAWABAN);
    var jData = jawabanSheet.getDataRange().getValues();
    for (var j = jData.length - 1; j >= 1; j--) {
      if (jData[j][0] && jData[j][0].toString().trim().toUpperCase() === token) {
        jawabanSheet.deleteRow(j + 1);
      }
    }

    var rowsJawaban = jawabanList.map(function(item) {
      return [token, item.idSoal, item.jawaban || '', item.nilai !== undefined ? item.nilai : 0, ts];
    });
    if (rowsJawaban.length) {
      jawabanSheet.getRange(jawabanSheet.getLastRow() + 1, 1, rowsJawaban.length, rowsJawaban[0].length).setValues(rowsJawaban);
    }

    // 2. Hitung skor per kode RIASEC
    var skor = { R: 0, I: 0, A: 0, S: 0, E: 0, C: 0 };

    var soalSheet = getSheet(SHEET_SOAL);
    var soalData = soalSheet.getDataRange().getValues();
    var soalMap = {};
    for (var s = 1; s < soalData.length; s++) {
      soalMap[soalData[s][0].toString()] = (soalData[s][2] || 'R').toString().trim().toUpperCase();
    }

    jawabanList.forEach(function(item) {
      var kode = soalMap[item.idSoal.toString()];
      if (kode && skor.hasOwnProperty(kode)) {
        skor[kode] += parseFloat(item.nilai) || 0;
      }
    });

    // 3. Urutkan ranking RIASEC
    var urutan = ['R', 'I', 'A', 'S', 'E', 'C'];
    var ranked = urutan.map(function(k) {
      return { kode: k, skor: skor[k], bidang: BIDANG_MAP[k] || k, namaInggris: BIDANG_INGGRIS[k] || k };
    }).sort(function(a, b) {
      return b.skor - a.skor;
    });

    var top3Kode = ranked.slice(0, 3).map(function(r) { return r.kode; }).join('');
    var interpretasi = getInterpretasiHolland(top3Kode);

    // 4. Ambil biodata terkini
    var biodata = cekBiodata(token);

    // 5. Simpan / Perbarui ke sheet Hasil
    var hasilSheet = getSheet(SHEET_HASIL);
    var hData = hasilSheet.getDataRange().getValues();
    var hasilRowIndex = -1;

    for (var h = 1; h < hData.length; h++) {
      if (hData[h][0] && hData[h][0].toString().trim().toUpperCase() === token) {
        hasilRowIndex = h + 1;
        break;
      }
    }

    var hasilValues = [
      token,
      biodata.nama || '',
      biodata.jenisKelamin || 'Laki-laki',
      biodata.tanggalLahirText || biodata.tanggalLahir || '',
      biodata.pendidikan || 'SD',
      biodata.sekolah || '',
      skor['R'], skor['I'], skor['A'], skor['S'], skor['E'], skor['C'],
      ranked[0].bidang,
      ranked[1] ? ranked[1].bidang : '',
      ranked[2] ? ranked[2].bidang : '',
      top3Kode,
      interpretasi.konsistensi,
      ts
    ];

    if (hasilRowIndex > 1) {
      hasilSheet.getRange(hasilRowIndex, 1, 1, hasilValues.length).setValues([hasilValues]);
    } else {
      hasilSheet.appendRow(hasilValues);
    }

    return {
      success      : true,
      token        : token,
      skor         : skor,
      ranked       : ranked,
      top3Kode     : top3Kode,
      biodata      : biodata,
      interpretasi : interpretasi,
      bidangUtama  : ranked[0].bidang,
      bidang2      : ranked[1] ? ranked[1].bidang : '',
      bidang3      : ranked[2] ? ranked[2].bidang : ''
    };
  } catch(err) {
    return { success: false, message: 'Gagal memproses jawaban: ' + err.message };
  }
}

// ============================================================
// CEK HASIL BY TOKEN
// ============================================================
function cekHasilByToken(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_HASIL);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;

    var col = getColumnMap(sheet);
    var cToken = col['token'] !== undefined ? col['token'] : 0;
    var cNama  = col['nama_siswa'] !== undefined ? col['nama_siswa'] : 1;
    var cJk    = col['jenis_kelamin'] !== undefined ? col['jenis_kelamin'] : -1;
    var cTgl   = col['tanggal_lahir'] !== undefined ? col['tanggal_lahir'] : -1;
    var cPend  = col['pendidikan'] !== undefined ? col['pendidikan'] : (col['jenjang'] !== undefined ? col['jenjang'] : 4);
    var cSek   = col['asal_sekolah'] !== undefined ? col['asal_sekolah'] : 5;
    var cR     = col['skor_r'] !== undefined ? col['skor_r'] : 6;
    var cI     = col['skor_i'] !== undefined ? col['skor_i'] : 7;
    var cA     = col['skor_a'] !== undefined ? col['skor_a'] : 8;
    var cS     = col['skor_s'] !== undefined ? col['skor_s'] : 9;
    var cE     = col['skor_e'] !== undefined ? col['skor_e'] : 10;
    var cC     = col['skor_c'] !== undefined ? col['skor_c'] : 11;
    var cUtama = col['bidang_utama'] !== undefined ? col['bidang_utama'] : 12;
    var cB2    = col['bidang_2'] !== undefined ? col['bidang_2'] : 13;
    var cB3    = col['bidang_3'] !== undefined ? col['bidang_3'] : 14;
    var cKod   = col['kode_holland'] !== undefined ? col['kode_holland'] : 15;
    var cKon   = col['konsistensi'] !== undefined ? col['konsistensi'] : 16;
    var cTs    = col['timestamp'] !== undefined ? col['timestamp'] : 17;

    for (var i = 1; i < data.length; i++) {
      if (data[i][cToken] && data[i][cToken].toString().trim().toUpperCase() === token) {
        var row = data[i];
        var skor = {
          R: parseFloat(row[cR]) || 0,
          I: parseFloat(row[cI]) || 0,
          A: parseFloat(row[cA]) || 0,
          S: parseFloat(row[cS]) || 0,
          E: parseFloat(row[cE]) || 0,
          C: parseFloat(row[cC]) || 0
        };

        var ranked = ['R', 'I', 'A', 'S', 'E', 'C'].map(function(k) {
          return { kode: k, skor: skor[k], bidang: BIDANG_MAP[k] || k, namaInggris: BIDANG_INGGRIS[k] || k };
        }).sort(function(a, b) { return b.skor - a.skor; });

        var top3Kode = (cKod >= 0 && row[cKod]) ? row[cKod].toString().trim() :
                       ranked.slice(0, 3).map(function(r) { return r.kode; }).join('');

        var interpretasi = getInterpretasiHolland(top3Kode);
        var bio = cekBiodata(token);

        return {
          token       : token,
          nama        : (bio.ada && bio.nama) ? bio.nama : (row[cNama] || token),
          jenisKelamin: (bio.ada && bio.jenisKelamin) ? bio.jenisKelamin : (cJk >= 0 ? (row[cJk] || 'Laki-laki') : 'Laki-laki'),
          tanggalLahir: (bio.ada && (bio.tanggalLahirText || bio.tanggalLahir)) ? (bio.tanggalLahirText || bio.tanggalLahir) : (cTgl >= 0 ? safeFormatDate(row[cTgl], 'dd MMM yyyy') : ''),
          pendidikan  : (bio.ada && bio.pendidikan) ? bio.pendidikan : (row[cPend] || 'SD'),
          jenjang     : (bio.ada && bio.pendidikan) ? bio.pendidikan : (row[cPend] || 'SD'),
          sekolah     : (bio.ada && bio.sekolah) ? bio.sekolah : (cSek >= 0 ? (row[cSek] || '') : ''),
          skor        : skor,
          ranked      : ranked,
          top3Kode    : top3Kode,
          interpretasi: interpretasi,
          bidangUtama : row[cUtama] || ranked[0].bidang,
          bidang2     : row[cB2] || (ranked[1] ? ranked[1].bidang : ''),
          bidang3     : row[cB3] || (ranked[2] ? ranked[2].bidang : ''),
          timestamp   : safeFormatDate(row[cTs], 'dd/MM/yyyy HH:mm')
        };
      }
    }
    return null;
  } catch(err) {
    return null;
  }
}

// ============================================================
// ADMIN: AMBIL SEMUA HASIL (LAPORAN) - ANTI CRASH
// ============================================================
function getAllHasil(filterJenjang, filterSekolah) {
  try {
    var sheet = getSheet(SHEET_HASIL);
    var data = sheet.getDataRange().getValues();
    var result = [];

    if (data.length <= 1) {
      return { success: true, data: [], total: 0 };
    }

    var col = getColumnMap(sheet);
    var cToken = col['token'] !== undefined ? col['token'] : 0;
    var cNama  = col['nama_siswa'] !== undefined ? col['nama_siswa'] : 1;
    var cJk    = col['jenis_kelamin'] !== undefined ? col['jenis_kelamin'] : -1;
    var cTgl   = col['tanggal_lahir'] !== undefined ? col['tanggal_lahir'] : -1;
    var cPend  = col['pendidikan'] !== undefined ? col['pendidikan'] : (col['jenjang'] !== undefined ? col['jenjang'] : 4);
    var cSek   = col['asal_sekolah'] !== undefined ? col['asal_sekolah'] : 5;
    var cR     = col['skor_r'] !== undefined ? col['skor_r'] : 6;
    var cI     = col['skor_i'] !== undefined ? col['skor_i'] : 7;
    var cA     = col['skor_a'] !== undefined ? col['skor_a'] : 8;
    var cS     = col['skor_s'] !== undefined ? col['skor_s'] : 9;
    var cE     = col['skor_e'] !== undefined ? col['skor_e'] : 10;
    var cC     = col['skor_c'] !== undefined ? col['skor_c'] : 11;
    var cUtama = col['bidang_utama'] !== undefined ? col['bidang_utama'] : 12;
    var cB2    = col['bidang_2'] !== undefined ? col['bidang_2'] : 13;
    var cB3    = col['bidang_3'] !== undefined ? col['bidang_3'] : 14;
    var cKod   = col['kode_holland'] !== undefined ? col['kode_holland'] : 15;
    var cKon   = col['konsistensi'] !== undefined ? col['konsistensi'] : 16;
    var cTs    = col['timestamp'] !== undefined ? col['timestamp'] : 17;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[cToken]) continue;

      var jenjang = (row[cPend] || '').toString().trim().toUpperCase();
      var sekolah = (cSek >= 0 && row[cSek]) ? row[cSek].toString().trim() : '';

      if (filterJenjang && filterJenjang !== 'semua' && jenjang !== filterJenjang.toUpperCase()) continue;
      if (filterSekolah && filterSekolah !== 'semua' && sekolah.toLowerCase().indexOf(filterSekolah.toLowerCase()) === -1) continue;

      var skor = {
        R: parseFloat(row[cR]) || 0,
        I: parseFloat(row[cI]) || 0,
        A: parseFloat(row[cA]) || 0,
        S: parseFloat(row[cS]) || 0,
        E: parseFloat(row[cE]) || 0,
        C: parseFloat(row[cC]) || 0
      };

      var ranked = ['R', 'I', 'A', 'S', 'E', 'C'].map(function(k) {
        return { kode: k, skor: skor[k], bidang: BIDANG_MAP[k] || k };
      }).sort(function(a, b) { return b.skor - a.skor; });

      var top3Kode = (cKod >= 0 && row[cKod]) ? row[cKod].toString().trim() :
                     ranked.slice(0, 3).map(function(r) { return r.kode; }).join('');

      var konsistensi = (cKon >= 0 && row[cKon]) ? row[cKon].toString().trim() :
                        ((REFERENSI_RIASEC.konsistensi && REFERENSI_RIASEC.konsistensi[top3Kode.slice(0, 2)]) || 'Sedang');

      result.push({
        no          : result.length + 1,
        token       : row[cToken],
        nama        : row[cNama] || row[cToken],
        jenisKelamin: cJk >= 0 ? (row[cJk] || 'Laki-laki') : 'Laki-laki',
        tanggalLahir: cTgl >= 0 ? safeFormatDate(row[cTgl], 'dd MMM yyyy') : '',
        jenjang     : row[cPend] || 'SD',
        pendidikan  : row[cPend] || 'SD',
        sekolah     : sekolah,
        skor        : skor,
        ranked      : ranked,
        top3Kode    : top3Kode,
        konsistensi : konsistensi,
        bidangUtama : row[cUtama] || ranked[0].bidang,
        bidang2     : row[cB2] || (ranked[1] ? ranked[1].bidang : ''),
        bidang3     : row[cB3] || (ranked[2] ? ranked[2].bidang : ''),
        timestamp   : safeFormatDate(row[cTs], 'dd/MM/yyyy HH:mm')
      });
    }

    return { success: true, data: result, total: result.length };
  } catch(err) {
    return { success: false, message: 'Gagal mengambil data laporan: ' + err.message, data: [] };
  }
}

// ============================================================
// LAPORAN LENGKAP PER SISWA
// ============================================================
function getLaporanSiswa(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var hasil = cekHasilByToken(token);
    if (!hasil) return { success: false, message: 'Data hasil tidak ditemukan untuk token ' + token };

    var biodata = cekBiodata(token);
    var MAKS_PER_DIMENSI = 18;
    var skor = hasil.skor || {};
    var persen = {};
    var urutan = ['R', 'I', 'A', 'S', 'E', 'C'];
    urutan.forEach(function(k) {
      var s = parseFloat(skor[k]) || 0;
      persen[k] = Math.round((s / MAKS_PER_DIMENSI) * 100);
    });

    // Detail jawaban per soal
    var jawabanSheet = getSheet(SHEET_JAWABAN);
    var jData = jawabanSheet.getDataRange().getValues();
    var detailJawaban = [];
    for (var j = 1; j < jData.length; j++) {
      if (jData[j][0] && jData[j][0].toString().trim().toUpperCase() === token) {
        detailJawaban.push({
          idSoal : jData[j][1],
          jawaban: jData[j][2],
          nilai  : jData[j][3]
        });
      }
    }

    // Map teks soal
    var soalSheet = getSheet(SHEET_SOAL);
    var sData = soalSheet.getDataRange().getValues();
    var soalMap = {};
    for (var s = 1; s < sData.length; s++) {
      soalMap[sData[s][0].toString()] = {
        teks  : sData[s][1],
        kode  : sData[s][2],
        bidang: BIDANG_MAP[sData[s][2]] || sData[s][2]
      };
    }

    var jawabanLengkap = detailJawaban.map(function(item) {
      var info = soalMap[item.idSoal.toString()] || {};
      return {
        idSoal : item.idSoal,
        teks   : info.teks || '',
        kode   : info.kode || '',
        bidang : info.bidang || '',
        jawaban: item.jawaban,
        nilai  : item.nilai
      };
    });

    var rekapDimensi = {};
    urutan.forEach(function(k) { rekapDimensi[k] = { ya: 0, tidak: 0, total: 0 }; });
    jawabanLengkap.forEach(function(item) {
      var k = (item.kode || '').toUpperCase();
      if (rekapDimensi[k]) {
        rekapDimensi[k].total++;
        if (parseFloat(item.nilai) > 0) rekapDimensi[k].ya++;
        else rekapDimensi[k].tidak++;
      }
    });

    return {
      success       : true,
      token         : token,
      biodata       : biodata.ada ? biodata : hasil,
      skor          : skor,
      persen        : persen,
      ranked        : hasil.ranked,
      top3Kode      : hasil.top3Kode,
      interpretasi  : hasil.interpretasi,
      bidangUtama   : hasil.bidangUtama,
      bidang2       : hasil.bidang2,
      bidang3       : hasil.bidang3,
      rekapDimensi  : rekapDimensi,
      jawabanLengkap: jawabanLengkap,
      timestamp     : hasil.timestamp
    };
  } catch(err) {
    return { success: false, message: 'Gagal mengambil laporan siswa: ' + err.message };
  }
}

// ============================================================
// ADMIN: STATISTIK DASHBOARD
// ============================================================
function getStatistik() {
  try {
    var tokenSheet = getSheet(SHEET_TOKEN);
    var hasilSheet = getSheet(SHEET_HASIL);
    var soalSheet  = getSheet(SHEET_SOAL);

    var tData = tokenSheet.getDataRange().getValues();
    var hData = hasilSheet.getDataRange().getValues();
    var sData = soalSheet.getDataRange().getValues();

    var totalToken = Math.max(tData.length - 1, 0);
    var tokenAktif = 0, tokenDipakai = 0;
    for (var i = 1; i < tData.length; i++) {
      var st = (tData[i][5] || '').toString().toLowerCase();
      if (st === 'aktif') tokenAktif++;
      else if (st === 'dipakai') tokenDipakai++;
    }

    var totalHasil = Math.max(hData.length - 1, 0);
    var totalSoal  = Math.max(sData.length - 1, 0);

    return {
      success     : true,
      totalToken  : totalToken,
      tokenAktif  : tokenAktif,
      tokenDipakai: tokenDipakai,
      totalHasil  : totalHasil,
      totalSoal   : totalSoal
    };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// ADMIN: TOKEN MANAGEMENT
// ============================================================
function generateToken(jumlah, jenjang, sekolah, namaWali) {
  try {
    var sheet = getSheet(SHEET_TOKEN);
    var tokens = [];
    var prefix = jenjang.toUpperCase() === 'SD' ? 'SD' : jenjang.toUpperCase() === 'TK' ? 'TK' : 'UM';

    var rows = [];
    for (var i = 0; i < jumlah; i++) {
      var token = prefix + '-' + generateRandomString(6);
      rows.push([token, namaWali || '', '', jenjang, sekolah, 'aktif', new Date(), '']);
      tokens.push(token);
    }
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, tokens: tokens, count: tokens.length };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function generateRandomString(length) {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var res = '';
  for (var i = 0; i < length; i++) {
    res += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return res;
}

function getAllToken(filterJenjang, filterStatus) {
  try {
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var jenjang = (row[3] || '').toString().toUpperCase();
      var status  = (row[5] || '').toString().toLowerCase();

      if (filterJenjang && filterJenjang !== 'semua' && jenjang !== filterJenjang.toUpperCase()) continue;
      if (filterStatus && filterStatus !== 'semua' && status !== filterStatus.toLowerCase()) continue;

      result.push({
        token      : row[0],
        namaWali   : row[1] || '',
        noHpWali   : row[2] || '',
        jenjang    : row[3] || '',
        asalSekolah: row[4] || '',
        status     : row[5] || 'aktif',
        tglDibuat  : safeFormatDate(row[6], 'dd/MM/yyyy HH:mm'),
        tglDipakai : safeFormatDate(row[7], 'dd/MM/yyyy HH:mm')
      });
    }
    return { success: true, data: result, total: result.length };
  } catch(err) {
    return { success: false, message: err.message, data: [] };
  }
}

function updateToken(token, dataUpdate) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        var rowIdx = i + 1;
        if (dataUpdate.jenjang     !== undefined) sheet.getRange(rowIdx, 4).setValue(dataUpdate.jenjang);
        if (dataUpdate.asalSekolah !== undefined) sheet.getRange(rowIdx, 5).setValue(dataUpdate.asalSekolah);
        if (dataUpdate.namaWali    !== undefined) sheet.getRange(rowIdx, 2).setValue(dataUpdate.namaWali);
        if (dataUpdate.noHpWali    !== undefined) sheet.getRange(rowIdx, 3).setValue(dataUpdate.noHpWali);
        if (dataUpdate.status      !== undefined) sheet.getRange(rowIdx, 6).setValue(dataUpdate.status);
        return { success: true, message: 'Token ' + token + ' berhasil diperbarui.' };
      }
    }
    return { success: false, message: 'Token tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function hapusToken(token) {
  return nonaktifkanToken(token);
}

function nonaktifkanToken(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        sheet.getRange(i + 1, 6).setValue('nonaktif');
        return { success: true, message: 'Token ' + token + ' berhasil dinonaktifkan.' };
      }
    }
    return { success: false, message: 'Token tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function aktifkanToken(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        sheet.getRange(i + 1, 6).setValue('aktif');
        return { success: true, message: 'Token ' + token + ' berhasil diaktifkan kembali.' };
      }
    }
    return { success: false, message: 'Token tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function hapusTokenPermanen(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Token ' + token + ' berhasil dihapus permanen.' };
      }
    }
    return { success: false, message: 'Token tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// ADMIN: SOAL MANAGEMENT & IMPORT
// ============================================================
function getAllSoal() {
  try {
    var sheet = getSheet(SHEET_SOAL);
    var data = sheet.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      result.push({
        id     : row[0],
        teks   : row[1],
        kode   : row[2],
        jenjang: row[3],
        opsiA  : row[4],
        opsiB  : row[5],
        nilaiA : row[6],
        nilaiB : row[7],
        aktif  : row[8]
      });
    }
    return { success: true, data: result, total: result.length };
  } catch(err) {
    return { success: false, message: err.message, data: [] };
  }
}

function importSoal(dataSoal, hapusLama) {
  try {
    var sheet = getSheet(SHEET_SOAL);
    if (hapusLama) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
    }

    var startId = sheet.getLastRow();
    var rows = dataSoal.map(function(s, idx) {
      var idStr = 'S' + ('000' + (startId + idx)).slice(-3);
      return [
        idStr,
        s.teks,
        s.kode ? s.kode.toUpperCase() : 'R',
        s.jenjang ? s.jenjang.toUpperCase() : 'SEMUA',
        s.opsiA || 'Ya',
        s.opsiB || 'Tidak',
        s.nilaiA !== undefined ? s.nilaiA : 1,
        s.nilaiB !== undefined ? s.nilaiB : 0,
        'aktif'
      ];
    });

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, message: 'Berhasil mengimpor ' + rows.length + ' soal.' };
  } catch(err) {
    return { success: false, message: 'Gagal impor soal: ' + err.message };
  }
}

function hapusSemuaSoal() {
  try {
    var sheet = getSheet(SHEET_SOAL);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return { success: true, message: 'Semua soal berhasil dihapus.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}
