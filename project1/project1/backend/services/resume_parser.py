import io
import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple

from PyPDF2 import PdfReader

try:
    import docx  # type: ignore
except ImportError:
    docx = None


class ResumeParser:
    """Lightweight resume parsing using PyPDF2 (and optional python-docx). No spaCy."""

    def __init__(self):
        self.taxonomy = self._get_massive_taxonomy()

    def _get_massive_taxonomy(self) -> Dict[str, List[str]]:
        return {
            "Programming Languages": [
                "Python", "Java", "C", "C++", "C#", "JavaScript", "TypeScript", "PHP", "Ruby", "Swift", "Kotlin", "Go",
                "Rust", "Scala", "Perl", "Haskell", "SQL", "R", "MATLAB", "Dart", "Solidity", "Assembly", "VB.NET", "Bash",
                "Cobol", "Fortran", "Lisp", "Prolog", "Objective-C", "Elixir", "Erlang", "F#", "Groovy", "Julia", "Lua", "PowerShell",
            ],
            "Web Frameworks & Libraries": [
                "React", "React.js", "Angular", "Vue", "Vue.js", "Next.js", "Nuxt.js", "Svelte", "jQuery", "Bootstrap",
                "Tailwind CSS", "Redux", "MobX", "Express", "Django", "Flask", "FastAPI", "Spring Boot", "Laravel",
                "Ruby on Rails", "Asp.net", "NestJS", "GraphQL", "Material UI", "Chakra UI", "SASS", "LESS", "Webpack",
                "Vite", "Babel", "Gulp", "Prisma", "Sequelize", "TypeORM", "Mongoose", "Socket.io", "Astro",
            ],
            "Industrial Automation and Control": [
                "PLC Programming", "Siemens PLC", "S7 300", "S7 1200", "S7-300", "S7-1200", "OMRON PLC", "Allen Bradley", "AB PLC", "ABB PLC",
                "SCADA", "HMI", "VFD", "PROFIBUS", "PROFINET", "ASI Communication", "Ladder Diagram",
                "Functional Block Diagram", "FBD", "WinCC", "Simatic Manager", "CX Programmer", "Keyence", "ASRS Warehouse",
                "Track and Trace", "SEW VFD", "Parameter Setting", "Encoder", "IR Device", "Profibus Coupler",
                "Control Panel Design", "Panel Wiring", "Electrical Load Calculation", "Cable Sizing", "SCADA Integration",
                "Industrial Network Troubleshooting", "Ethernet/IP", "Modbus RTU", "Modbus TCP", "Safety PLC",
                "Emergency System Integration", "Servo Motor", "Motion Control", "PID Control Tuning", "Root Cause Analysis",
                "RCA", "Preventive Maintenance", "Predictive Maintenance", "Energy Management Systems",
            ],
            "Industry 4.0 and Smart Manufacturing": [
                "IIoT", "Industrial Internet of Things", "Digital Twin", "MES", "Manufacturing Execution Systems",
                "OEE", "Overall Equipment Effectiveness", "Data Logging", "Process Optimization", "Remote Monitoring",
                "Barcode System", "RFID Integration", "Industrial Cybersecurity",
            ],
            "Engineering and Maintenance": [
                "Equipment Commissioning", "Site Maintenance", "Troubleshooting", "Production Control",
                "Material Management", "Procurement", "Client Relationship Management", "Manpower Management",
                "AutoCAD", "EPLAN", "Circuit Breaker", "Bus Bar", "Induction Motor", "Sensors", "Installation",
                "Electrical Installation", "PLC Projects", "Wehrhann AAC Plant", "Contactor", "Relay",
                "Preventive Maintenance", "Corrective Maintenance", "Calibration", "P&ID", "Schematic Reading",
                "BOM Management", "Spare Parts Management", "Quality Control", "Lean Manufacturing Principles",
            ],
            "IT & Networking": [
                "Data Center", "Server Installation", "Storage System", "Network Switches", "Firewall Configuration",
                "RedHat", "CentOS", "Oracle SQL", "MSSQL", "Tomcat", "Web Application Installation",
                "Virtualization", "VMware", "Hyper-V", "Backup & Recovery", "Disaster Recovery",
                "Network Architecture", "Cloud Basics", "AWS Fundamentals", "Azure Fundamentals",
                "Linux Server Hardening", "Database Optimization", "API Integration", "Active Directory", "DNS", "DHCP",
                "VPN Configuration", "Network Monitoring", "Wi-Fi Networks", "VoIP", "SAN", "NAS", "Load Balancing",
            ],
            "Project & Business Management": [
                "Project Planning", "Project Scheduling", "MS Project", "Primavera", "Budget Planning",
                "Cost Control", "Vendor Evaluation", "Technical Negotiation", "Risk Assessment",
                "Mitigation Planning", "KPI Development", "Performance Tracking", "SOP Development",
                "Lean Manufacturing", "Six Sigma", "Kaizen", "Continuous Improvement", "Stakeholder Management",
                "Resource Allocation", "Change Management", "Quality Management", "Procurement Management",
                "Contract Management", "Business Process Improvement", "Strategic Planning", "Market Analysis",
            ],
            "Leadership & Professional": [
                "Leadership", "Decision-Making", "Conflict Resolution", "Cross-Functional Coordination",
                "Technical Training", "Mentoring", "Stakeholder Management", "Change Management",
                "Time Management", "Priority Management", "Presentation Skills", "Reporting Skills",
                "Motivational Leader", "Strategic Thinker", "Team Player", "Innovative", "Communicator",
                "Coaching", "Delegation", "Performance Management", "Emotional Intelligence", "Adaptability",
                "Problem Solving", "Critical Thinking", "Negotiation", "Public Speaking", "Active Listening",
            ],
            "Certifications": [
                "PMP", "Project Management Professional", "Six Sigma Green Belt", "Six Sigma Black Belt",
                "Certified Automation Professional", "CAP", "ITIL Foundation", "CCNA", "Industrial Networking",
                "CompTIA A+", "CompTIA Network+", "CompTIA Security+", "AWS Certified Solutions Architect",
                "Azure Administrator Associate", "Certified Ethical Hacker", "CEH", "Scrum Master", "CSM",
            ],
            "Cybersecurity": [
                "Ethical Hacking", "Penetration Testing", "Network Security", "Information Security", "Cryptography",
                "Firewalls", "VPN", "SIEM", "Vulnerability Assessment", "Incident Response", "OWASP", "Metasploit",
                "Wireshark", "Burp Suite", "Malware Analysis", "Security Audits", "Compliance (GDPR, HIPAA)",
                "Risk Management", "Identity and Access Management", "IAM", "Endpoint Security", "Intrusion Detection Systems",
                "IDS", "Intrusion Prevention Systems", "IPS", "Security Information and Event Management",
            ],
        }

    def _clean_text(self, text: str) -> str:
        text = re.sub(r"[^\x00-\x7F]+", " ", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def extract_text(self, file_content: bytes, filename: str) -> str:
        text = ""
        lower = filename.lower()
        if lower.endswith(".pdf"):
            try:
                reader = PdfReader(io.BytesIO(file_content))
                for page in reader.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
            except Exception as e:
                print(f"Error extracting PDF: {e}")
            return self._clean_text(text)
        if lower.endswith((".docx", ".doc")):
            if docx is None:
                print("python-docx not installed; DOCX text extraction skipped")
                return ""
            try:
                document = docx.Document(io.BytesIO(file_content))
                return self._clean_text("\n".join(p.text for p in document.paragraphs))
            except Exception as e:
                print(f"Error extracting DOCX: {e}")
                return ""
        return ""

    def _skill_in_text(self, skill: str, text_lower: str) -> bool:
        if len(skill.strip()) < 2:
            return False
        s = skill.lower()
        if len(s) <= 3 and s in ("c", "r", "go"):
            return bool(re.search(rf"(?<![a-z0-9]){re.escape(s)}(?![a-z0-9])", text_lower))
        return s in text_lower

    def _match_taxonomy(self, text: str) -> Tuple[Dict[str, Set[str]], Set[str]]:
        text_lower = text.lower()
        categorized: Dict[str, Set[str]] = {cat: set() for cat in self.taxonomy}
        all_found: Set[str] = set()
        for category, skills in self.taxonomy.items():
            flat = sorted(skills, key=len, reverse=True)
            for skill in flat:
                if self._skill_in_text(skill, text_lower):
                    categorized[category].add(skill)
                    all_found.add(skill)
        return categorized, all_found

    def _extract_soft_skills_keywords(self, text_lower: str) -> Set[str]:
        out: Set[str] = set()
        soft_skill_keywords = {
            "Leadership": ["lead ", "managed", "supervised", "directed", "headed", "mentored"],
            "Communication": ["presented", "communicated", "wrote", "spoke", "negotiated"],
            "Teamwork": ["collaborated", "coordinated", "team ", "assisted", "helped"],
            "Problem Solving": ["solved", "analyzed", "resolved", "fixed", "optimized", "troubleshot"],
        }
        for skill, indicators in soft_skill_keywords.items():
            if any(ind in text_lower for ind in indicators):
                out.add(skill)
        return out

    def _extract_contact_info(self, text: str) -> Dict[str, Optional[str]]:
        email = re.findall(r"[\w.-]+@[\w.-]+\.\w+", text)
        phone = re.findall(r"\(?\+?\d{1,4}?[\s.-]?\d{3,4}[\s.-]?\d{3,4}[\s.-]?\d{3,9}", text)
        linkedin = re.findall(r"linkedin\.com/in/[\w-]+", text)
        github = re.findall(r"github\.com/[\w-]+", text)
        return {
            "email": email[0] if email else None,
            "phone": phone[0] if phone else None,
            "linkedin": f"https://{linkedin[0]}" if linkedin else None,
            "github": f"https://{github[0]}" if github else None,
        }

    def _extract_education_details(self, text: str) -> List[Dict[str, Any]]:
        edu_list: List[Dict[str, Any]] = []
        for line in text.split("\n"):
            if any(
                key in line.upper()
                for key in [
                    "B.TECH",
                    "B.SC",
                    "B.E",
                    "M.TECH",
                    "MS",
                    "PHD",
                    "BACHELOR",
                    "MASTER",
                    "UNIVERSITY",
                    "COLLEGE",
                    "INSTITUTE",
                ]
            ):
                year_match = re.search(r"(20\d{2}|19\d{2})", line)
                edu_list.append({"raw": line.strip(), "year": year_match.group(1) if year_match else None})
        return edu_list

    def parse_resume(self, file_content: bytes, filename: str) -> Dict[str, Any]:
        if isinstance(file_content, str):
            text = self._clean_text(file_content)
        else:
            text = self.extract_text(file_content, filename)

        if not text.strip():
            return {"error": "Document appears empty or unreadable. For DOCX, install python-docx: pip install python-docx"}

        text_lower = text.lower()
        categorized_skills, all_found = self._match_taxonomy(text)
        contextual_soft = self._extract_soft_skills_keywords(text_lower)
        if contextual_soft:
            categorized_skills.setdefault("Soft Skills", set())
            categorized_skills["Soft Skills"].update(contextual_soft)
            all_found.update(contextual_soft)

        contact_info = self._extract_contact_info(text)
        education = self._extract_education_details(text)

        cgpa = None
        percentage = None
        cgpa_match = re.search(
            r"(?:CGPA|GPA|SGPA|Pointer)[\s:]*([0-9]\.?[0-9]{0,2})(?:/[1]?[0])?",
            text,
            re.IGNORECASE,
        )
        if cgpa_match:
            cgpa = cgpa_match.group(1)
        else:
            edu_text = "\n".join(
                line for line in text.split("\n") if any(x in line.upper() for x in ("EDUCATION", "UNIVERSITY", "COLLEGE"))
            )
            cgpa_fallback = re.search(r"\b([4-9]\.[0-9]{1,2}|10\.0)\b", edu_text or text)
            if cgpa_fallback:
                cgpa = cgpa_fallback.group(1)

        perc_match = re.search(r"(\d{2}(?:\.\d{1,2})?)[\s]*(?:%|percent|percentage)", text, re.IGNORECASE)
        if perc_match:
            percentage = perc_match.group(1)

        exp_match = re.search(r"(\d+)\+?\s*(?:years|yrs|year)", text, re.IGNORECASE)
        experience = exp_match.group(1) if exp_match else "0"

        final_skills: Dict[str, List[str]] = {}
        for cat, skills in categorized_skills.items():
            if skills:
                final_skills[cat] = sorted(skills, key=lambda x: x.lower())

        flat_skills: List[str] = []
        for s_list in final_skills.values():
            flat_skills.extend(s_list)
        flat_skills = sorted(set(flat_skills), key=lambda x: x.lower())

        score = 40
        score += min(40, len(all_found) * 5)
        if cgpa or percentage:
            score += 5
        if int(experience) > 0:
            score += 5

        if cgpa:
            academic_score = cgpa
            score_type = "CGPA"
        elif percentage:
            academic_score = f"{percentage}%"
            score_type = "Percentage"
        else:
            academic_score = "N/A"
            score_type = "Score"

        return {
            "skills": flat_skills,
            "categorized_skills": final_skills,
            "education": education,
            "contact": contact_info,
            "cgpa": cgpa,
            "percentage": percentage,
            "academic_score": academic_score,
            "score_type": score_type,
            "years_of_experience": experience,
            "ats_score": min(100, score),
            "status": "success",
            "filename": filename,
            "text_preview": text[:300].replace("\n", " ").strip() + "...",
        }


resume_parser = ResumeParser()
