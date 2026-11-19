import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_path):
    prs = Presentation()
    # 16:9 Widescreen: 13.333 x 7.5 inches
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # Blank slide

    # Theme Colors
    C_NAVY_DARK = RGBColor(15, 23, 42)     # #0F172A
    C_NAVY_CARD = RGBColor(30, 41, 59)     # #1E293B
    C_PRIMARY   = RGBColor(2, 132, 199)    # #0284C7
    C_PRIMARY_DK= RGBColor(3, 105, 161)    # #0369A1
    C_SKY_LIGHT = RGBColor(224, 242, 254)  # #E0F2FE
    C_CARD_BG   = RGBColor(248, 250, 252)  # #F8FAFC
    C_CARD_PRI  = RGBColor(240, 249, 255)  # #F0F9FF
    C_BORDER    = RGBColor(226, 232, 240)  # #E2E8F0
    C_BORDER_PRI= RGBColor(186, 230, 253)  # #BAE6FD
    C_TEXT_DARK = RGBColor(15, 23, 42)     # #0F172A
    C_TEXT_BODY = RGBColor(51, 65, 85)     # #334155
    C_TEXT_MUTED= RGBColor(100, 116, 139)  # #64748B
    C_WHITE     = RGBColor(255, 255, 255)  # #FFFFFF
    C_EMERALD   = RGBColor(16, 185, 129)   # #10B981
    C_AMBER     = RGBColor(245, 158, 11)   # #F59E0B
    C_RED       = RGBColor(239, 68, 68)    # #EF4444

    def add_header(slide, category, title, slide_num, total_slides=18):
        # Top banner category
        tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8.5), Inches(0.35))
        tf_cat = tb_cat.text_frame
        tf_cat.word_wrap = True
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = C_PRIMARY

        # Slide Title
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(8.5), Inches(0.65))
        tf_title = tb_title.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = C_TEXT_DARK

        # Department / Institution Tag (Right aligned)
        tb_inst = slide.shapes.add_textbox(Inches(9.5), Inches(0.4), Inches(3.0), Inches(0.8))
        tf_inst = tb_inst.text_frame
        tf_inst.word_wrap = True
        p_inst1 = tf_inst.paragraphs[0]
        p_inst1.alignment = PP_ALIGN.RIGHT
        p_inst1.text = "TKIET Warananagar"
        p_inst1.font.size = Pt(11)
        p_inst1.font.bold = True
        p_inst1.font.color.rgb = C_TEXT_MUTED

        p_inst2 = tf_inst.add_paragraph()
        p_inst2.alignment = PP_ALIGN.RIGHT
        p_inst2.text = "SY MCA • Review-II (2026–27)"
        p_inst2.font.size = Pt(10)
        p_inst2.font.color.rgb = C_PRIMARY_DK

        # Header divider line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.733), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = C_BORDER
        line.line.color.rgb = C_BORDER

        # Footer
        tb_foot = slide.shapes.add_textbox(Inches(0.8), Inches(7.0), Inches(11.733), Inches(0.35))
        tf_foot = tb_foot.text_frame
        p_foot = tf_foot.paragraphs[0]
        p_foot.text = f"FactoryMind AI: Predictive Manufacturing Platform  |  Department of Computer Applications, TKIET  |  Slide {slide_num} of {total_slides}"
        p_foot.font.size = Pt(9.5)
        p_foot.font.color.rgb = C_TEXT_MUTED

    def add_card(slide, x, y, w, h, title, bullets, is_primary=False, badge=None):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = C_CARD_PRI if is_primary else C_CARD_BG
        shape.line.color.rgb = C_BORDER_PRI if is_primary else C_BORDER
        shape.line.width = Pt(1.5 if is_primary else 1)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.25)
        tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.25)
        tf.margin_bottom = Inches(0.25)

        # Title
        p_title = tf.paragraphs[0]
        header_text = f"● {title.upper()}" if not badge else f"[{badge}] {title.upper()}"
        p_title.text = header_text
        p_title.font.size = Pt(13)
        p_title.font.bold = True
        p_title.font.color.rgb = C_PRIMARY_DK if is_primary else C_PRIMARY
        p_title.space_after = Pt(8)

        # Bullets
        for b in bullets:
            p = tf.add_paragraph()
            p.text = f"•  {b}"
            p.font.size = Pt(11)
            p.font.color.rgb = C_TEXT_BODY
            p.space_after = Pt(6)

    slides_content = [
        # SLIDE 1: Cover Page
        {
            "is_cover": True,
            "title": "FactoryMind AI",
            "subtitle": "An AI-Based Predictive Manufacturing and Machine Health Monitoring Platform",
            "category": "SY MCA Major Project — Preliminary Project Report (Second Presentation)",
            "institution": "Tatyasaheb Kore Institute of Engineering & Technology (TKIET), Warananagar\nDepartment of Computer Applications • Academic Year: 2026–2027",
            "members": [
                "1. [Student Name 1] — Roll No: [Roll 1]",
                "2. [Student Name 2] — Roll No: [Roll 2]",
                "3. [Student Name 3] — Roll No: [Roll 3]"
            ],
            "guide": [
                "[Name of Project Guide]",
                "Assistant Professor, Department of Computer Applications",
                "TKIET, Warananagar"
            ]
        },
        # SLIDE 2: Abstract
        {
            "category": "Executive Overview",
            "title": "Abstract",
            "cards": [
                {
                    "title": "Problem Statement & Solution",
                    "is_primary": True,
                    "bullets": [
                        "Problem: Unplanned machine breakdowns cost manufacturing plants $260,000/hour in emergency downtime.",
                        "Solution: FactoryMind AI is an end-to-end cyber-physical predictive maintenance platform integrating IIoT telemetry, Machine Learning, and Explainable AI (XAI).",
                        "Continuous Telemetry: Monitors temperature, vibration RMS, current draw, and spindle RPM across industrial machines.",
                        "Predictive Paradigm: Replaces reactive repairs with algorithmic Remaining Useful Life (RUL) estimation and root-cause alerts."
                    ]
                },
                {
                    "title": "Technology & Expected Outcomes",
                    "is_primary": False,
                    "bullets": [
                        "Dual Architecture: High-throughput Java 21 Edge Gateway (PLC polling) + Python 3.12 FastAPI ML Inference.",
                        "Interactive 3D Digital Twin: Three.js WebGL visualization offering real-time plant floor thermal & vibration mapping.",
                        "Target Metrics: >93% machine failure classification accuracy and 30-40% reduction in unplanned factory downtime.",
                        "Academic & Industry Contribution: Fully functional open-standard cyber-physical platform meeting SY MCA Review-II requirements."
                    ]
                }
            ]
        },
        # SLIDE 3: Chapter 1 - Introduction
        {
            "category": "Chapter 1: Introduction",
            "title": "Project Background & Industrial Context",
            "cards": [
                {
                    "title": "Manufacturing Paradigm Shift",
                    "is_primary": True,
                    "bullets": [
                        "Industry 4.0 Transformation: Convergence of cyber-physical systems, IIoT, and cloud computing across smart factories.",
                        "Critical Industrial Assets: CNC lathes, hydraulic pumps, industrial turbines, and robotic arms undergo continuous mechanical stress.",
                        "Downtime Impact: Machine failure halts entire production lines, causing cascading financial losses and safety hazards.",
                        "Condition-Based Monitoring (CBM): Sensor data allows automated early detection of thermal expansion, bearing wear, and imbalance."
                    ]
                },
                {
                    "title": "Core Problem Addressed",
                    "is_primary": False,
                    "bullets": [
                        "The Reactive Maintenance Trap: 'Run-to-failure' creates catastrophic damages, high spare-part costs, and production delays.",
                        "Inefficient Preventive Overhauls: Calendar-based maintenance replaces healthy parts prematurely, wasting 35% of component lifespan.",
                        "Data Silos: Sensor data in traditional legacy PLCs is isolated in proprietary formats without predictive intelligence.",
                        "Need for FactoryMind AI: Bridges physical factory machines with intelligent AI diagnostics and automated work orders."
                    ]
                }
            ]
        },
        # SLIDE 4: Problem Statement & Significance
        {
            "category": "Chapter 1: Introduction (Contd.)",
            "title": "Problem Statement & Need of the Project",
            "cards": [
                {
                    "title": "Formal Problem Statement",
                    "is_primary": True,
                    "bullets": [
                        "Design and implement a real-time, low-latency predictive maintenance and digital twin platform.",
                        "Ingest multi-sensor industrial telemetry (vibration, thermal, electrical, rotational) via high-throughput edge gateways.",
                        "Execute Machine Learning models (Random Forest, Gradient Boosting) to classify anomalies before physical breakdown.",
                        "Calculate Remaining Useful Life (RUL) and generate automated dispatch orders to technicians."
                    ]
                },
                {
                    "title": "Need & Significance for Industry",
                    "is_primary": False,
                    "bullets": [
                        "Economic Significance: Reduces unplanned factory downtime by 30-40% and cuts maintenance overhead by 25%.",
                        "SME Accessibility: Existing enterprise solutions (Siemens MindSphere, GE Predix) cost $50K+/yr. FactoryMind AI delivers an open, affordable alternative.",
                        "Operator Safety: Early warnings on bearing failure and overheating prevent catastrophic workplace explosions and fires.",
                        "Asset Longevity: Prolongs industrial machinery lifespan through timely lubrication and mechanical balancing."
                    ]
                }
            ]
        },
        # SLIDE 5: Objectives & Scope
        {
            "category": "Chapter 1: Introduction (Contd.)",
            "title": "Project Objectives & Scope",
            "cards": [
                {
                    "title": "SMART Project Objectives",
                    "is_primary": True,
                    "bullets": [
                        "Objective 1: Build a high-throughput Java Edge Gateway to ingest simulated PLC telemetry at sub-second intervals.",
                        "Objective 2: Develop a Python FastAPI backend providing REST and WebSocket APIs for telemetry, predictions, and alerts.",
                        "Objective 3: Train Random Forest and Gradient Boosting models achieving >93% accuracy on NASA/Kaggle predictive datasets.",
                        "Objective 4: Construct a real-time 3D WebGL Digital Twin console rendering live machine status, heatmaps, and audio alerts."
                    ]
                },
                {
                    "title": "Scope & Project Boundaries",
                    "is_primary": False,
                    "bullets": [
                        "Target Machinery Scope: Milling machines, CNC lathes, centrifugal pumps, and hydraulic compressors.",
                        "Monitored Telemetry Scope: Spindle temperature (°C), vibration RMS (mm/s), motor current (A), and rotation speed (RPM).",
                        "In-Scope Deliverables: Edge gateway, ML inference pipeline, XAI explainability, 3D WebGL dashboard, and automated alerts.",
                        "Out-of-Scope (Future Scope): Hardware physical PLC installation and direct robotic emergency trip actuator integration."
                    ]
                }
            ]
        },
        # SLIDE 6: Target Users & Beneficiaries
        {
            "category": "Chapter 1: Introduction (Contd.)",
            "title": "Target Users & Stakeholder Beneficiaries",
            "cards": [
                {
                    "title": "Primary System Users",
                    "is_primary": False,
                    "bullets": [
                        "Plant Operations Managers: High-level plant overview, machine health status, overall equipment efficiency (OEE), and cost tracking.",
                        "Maintenance Technicians: Work orders, exact machine diagnostic insights, component failure reasons, and recommended action steps.",
                        "Reliability & Reliability Engineers: Sensor telemetry trends, vibration spectrum analytics, and model confidence scores.",
                        "System Administrators: User authentication, machine provisioning, role-based access control, and edge gateway status."
                    ]
                },
                {
                    "title": "Key Organizational Benefits",
                    "is_primary": True,
                    "bullets": [
                        "Zero Surprise Failures: Predictive warnings 48 to 72 hours prior to threshold breaches allow planned weekend repairs.",
                        "Optimized Spare Inventory: Eliminates emergency overnight part orders by ordering components based on RUL projections.",
                        "Audit & Compliance Trail: Complete ISO 55000 asset management logs with timestamped machine telemetry history.",
                        "Seamless Experience: Responsive single-page application with standalone login portal and instant live updates."
                    ]
                }
            ]
        },
        # SLIDE 7: Chapter 2 - Literature Survey
        {
            "category": "Chapter 2: Literature Survey",
            "title": "Literature Survey & Existing System Analysis",
            "cards": [
                {
                    "title": "Review of Key Research Papers",
                    "is_primary": True,
                    "bullets": [
                        "Paper 1 (Jardine et al., Mech. Systems & Signal Proc.): Foundational survey on CBM and machine health prognostics using vibration signal processing.",
                        "Paper 2 (Lee et al., IEEE Trans. Ind. Informatics): Industrial AI architecture for cyber-physical manufacturing and digital twin synchronization.",
                        "Paper 3 (NASA PHM Turbofan Dataset - Saxena et al.): Benchmark run-to-failure degradation modeling validating ensemble regression for RUL.",
                        "Key Finding: Machine learning combined with feature engineering (RMS, kurtosis) outperforms static threshold monitoring."
                    ]
                },
                {
                    "title": "Existing Commercial Systems",
                    "is_primary": False,
                    "bullets": [
                        "Siemens MindSphere: Enterprise IIoT cloud platform offering extensive hardware connectivity but expensive licensing.",
                        "GE Digital Predix: Industrial APM platform with advanced analytics tailored to heavy aviation and power generation.",
                        "PTC ThingWorx: Rapid IoT dashboard application builder requiring extensive proprietary add-ons.",
                        "Limitations: High licensing fees ($50K+/yr), complex proprietary lock-in, and lack of lightweight interactive 3D digital twins for SMEs."
                    ]
                }
            ]
        },
        # SLIDE 8: Research Gap Analysis
        {
            "category": "Chapter 2: Literature Survey (Contd.)",
            "title": "Research Gap & How Proposed System Solves It",
            "cards": [
                {
                    "title": "Gaps Identified in Existing Work",
                    "is_primary": False,
                    "bullets": [
                        "Gap 1 (High Cost Barrier): Existing enterprise solutions are financially inaccessible to medium and small manufacturing units.",
                        "Gap 2 (Black-Box AI): Commercial platforms predict failures without explaining root cause features (e.g. vibration vs temperature).",
                        "Gap 3 (Rigid Monolithic Stack): Difficult to run hybrid language workloads (Java edge PLC ingest + Python ML data science).",
                        "Gap 4 (Static 2D Visualizations): Legacy systems use tabular grids instead of interactive spatial 3D plant floor digital twins."
                    ]
                },
                {
                    "title": "FactoryMind AI Proposed Innovations",
                    "is_primary": True,
                    "bullets": [
                        "Affordable Open-Standard Architecture: Built on Python FastAPI, SQLite/PostgreSQL, and lightweight modern web standards.",
                        "Explainable AI (XAI): Feature importance metrics and transparent root-cause failure explanations with every alert.",
                        "Dual-Engine Synergy: Java 21 for deterministic edge gateway data acquisition + Python 3.12 for ML diagnostics.",
                        "Interactive 3D Digital Twin: Real-time Three.js spatial canvas rendering live machine health and thermal state."
                    ]
                }
            ]
        },
        # SLIDE 9: Chapter 3 - Proposed System
        {
            "category": "Chapter 3: Proposed System & Requirements",
            "title": "Proposed System & Key Innovations",
            "cards": [
                {
                    "title": "System Architecture Overview",
                    "is_primary": True,
                    "bullets": [
                        "Tier 1 - Industrial Edge: Java Edge Gateway pulls sensor telemetry from machine PLCs and publishes JSON payloads.",
                        "Tier 2 - AI Diagnostic Core: Python FastAPI processes telemetry, executes ML inference, and computes health index (0-100%).",
                        "Tier 3 - Persistence & Alerting: Relational schema logs sensor history, triggers alerts, and assigns maintenance tickets.",
                        "Tier 4 - Presentation Layer: Modern SPA with 3D WebGL Digital Twin, real-time Chart.js telemetry, and synthesized cyber audio."
                    ]
                },
                {
                    "title": "Key Functional Features",
                    "is_primary": False,
                    "bullets": [
                        "Real-time Telemetry Dashboard: Sub-second live sensor graphs for Vibration RMS, Temperature, Current, and RPM.",
                        "Predictive Health & RUL Index: Continuous scoring of machine condition with remaining operational hour forecasts.",
                        "Multi-Level Alert Engine: Automatic categorization of Warning, Critical, and Emergency states with audio sirens.",
                        "Automated Maintenance Dispatch: One-click work order creation and assignment to available technicians."
                    ]
                }
            ]
        },
        # SLIDE 10: Requirements Specification
        {
            "category": "Chapter 3: Proposed System & Requirements",
            "title": "Hardware, Software & System Requirements",
            "cards": [
                {
                    "title": "Hardware Specifications",
                    "is_primary": False,
                    "bullets": [
                        "Edge Gateway Node: Raspberry Pi 4 / Intel Celeron N5105, 4 GB RAM, Ethernet/WiFi (Simulated via Java Edge Daemon).",
                        "AI Backend Server: 4 Core CPU (Intel i5/AMD Ryzen 5), 8 GB RAM, 256 GB SSD.",
                        "Client Workstation: Standard PC/Laptop with WebGL-enabled GPU (Intel Iris/Nvidia GTX), modern Chrome/Edge browser.",
                        "Physical Sensors (Simulated): Piezoelectric vibration accelerometers, PT100 thermal probes, CT current sensors."
                    ]
                },
                {
                    "title": "Software & Functional Matrix",
                    "is_primary": True,
                    "bullets": [
                        "Operating System: Windows 10/11 / Ubuntu Linux 22.04 LTS.",
                        "Languages & Runtimes: Python 3.12, Java 21 LTS (OpenJDK), JavaScript (ES6+).",
                        "Frameworks & Libraries: FastAPI, Uvicorn, SQLAlchemy, Pydantic, Three.js, Chart.js, scikit-learn.",
                        "Non-Functional Targets: <200ms API latency, 99.9% uptime, responsive design, JWT token security, role-based authorization."
                    ]
                }
            ]
        },
        # SLIDE 11: Chapter 4 - System Architecture
        {
            "category": "Chapter 4: System Architecture & Design",
            "title": "System Architecture & Data Workflow",
            "cards": [
                {
                    "title": "4-Tier Architecture Pipeline",
                    "is_primary": True,
                    "bullets": [
                        "[Tier 1] IoT Telemetry Layer: High-frequency telemetry generation and edge buffering (Java Edge Gateway & Python Simulator).",
                        "[Tier 2] Communication Layer: RESTful JSON APIs and bidirectional WebSocket pipelines streaming 2.5s updates.",
                        "[Tier 3] AI Diagnostic Engine: Preprocessing, normalization, Random Forest classifier, and RUL regression models.",
                        "[Tier 4] Web Client Console: 3D Digital Twin, live sensor telemetry visualizer, maintenance tracker, and AI copilot."
                    ]
                },
                {
                    "title": "Data Flow Lifecycle",
                    "is_primary": False,
                    "bullets": [
                        "Step 1: Sensors emit vibration RMS (mm/s), temperature (°C), current (A), and RPM.",
                        "Step 2: Edge gateway batches readings and sends POST requests to /api/v1/telemetry endpoint.",
                        "Step 3: ML inference evaluates anomaly risk and updates machine Health Score (0-100%).",
                        "Step 4: If Health Score < 70%, automated alert created and broadcast to frontend with synthesized alert chime.",
                        "Step 5: Maintenance technician accepts work order and schedules corrective overhaul."
                    ]
                }
            ]
        },
        # SLIDE 12: UML & Design Diagrams
        {
            "category": "Chapter 4: System Architecture & Design",
            "title": "UML Design: Use Case & Database ER Schema",
            "cards": [
                {
                    "title": "Mandatory Use Case Summary",
                    "is_primary": True,
                    "bullets": [
                        "Plant Manager Actor: View plant aggregate health, inspect OEE metrics, review maintenance expenditures.",
                        "Maintenance Technician Actor: Acknowledge alerts, view diagnostic breakdown, complete work orders.",
                        "Reliability Engineer Actor: Query historical telemetry, inspect ML feature importance, recalibrate models.",
                        "Edge Gateway System Actor: Authenticate via API token, transmit real-time telemetry streams, log heartbeat."
                    ]
                },
                {
                    "title": "Database Schema Design (Relational ER)",
                    "is_primary": False,
                    "bullets": [
                        "Table 'users': Stores user credentials, hashed passwords (bcrypt), and roles (admin, operator, technician).",
                        "Table 'machines': Asset registry containing serial number, model, installation date, and operational status.",
                        "Table 'telemetry': High-frequency time-series table logging temperature, vibration, current, and RPM.",
                        "Table 'predictions': Model inference records, anomaly classifications, RUL forecasts, and feature weights.",
                        "Table 'maintenance_orders': Work orders linked to machines and assigned technician IDs."
                    ]
                }
            ]
        },
        # SLIDE 13: Chapter 5 - Tech Stack
        {
            "category": "Chapter 5: Modules & Technology Stack",
            "title": "Technology Stack & Modular Allocation",
            "cards": [
                {
                    "title": "Technology Stack & Justification",
                    "is_primary": True,
                    "bullets": [
                        "Backend - Python 3.12 & FastAPI: Asynchronous performance, automatic OpenAPI documentation, native ML interoperability.",
                        "Edge Layer - Java 21 LTS: Robust multi-threaded socket handling, industrial stability, and strict typing for edge devices.",
                        "Frontend - HTML5, CSS3, ES6 JavaScript: Zero-build instant loading, Three.js WebGL 3D graphics, and Chart.js telemetry charts.",
                        "Database - SQLite 3 / PostgreSQL: ACID-compliant time-series storage with SQLAlchemy ORM and Pydantic validation.",
                        "Machine Learning - Scikit-Learn: Random Forest & Gradient Boosting models optimized for low-latency inference (<15ms)."
                    ]
                },
                {
                    "title": "Modular Group Allocation",
                    "is_primary": False,
                    "bullets": [
                        "Module 1 (Edge Ingestion & Java Gateway): Group Member 1 — PLC socket listener, payload serialization, network reconnect.",
                        "Module 2 (FastAPI Core & Database ORM): Group Member 2 — User auth, REST APIs, WebSocket streaming, schema migrations.",
                        "Module 3 (AI/ML Diagnostics & RUL): Group Member 1 — Dataset preprocessing, model training, feature importance scoring.",
                        "Module 4 (3D WebGL Twin & UI Dashboard): Group Member 3 — Three.js rendering, Chart.js telemetry, audio sound effects."
                    ]
                }
            ]
        },
        # SLIDE 14: Chapter 6 - Implementation Progress
        {
            "category": "Chapter 6: Initial Implementation & Progress",
            "title": "Current Implementation Status (Review-II)",
            "cards": [
                {
                    "title": "Completed Modules (100% Functional)",
                    "is_primary": True,
                    "bullets": [
                        "✓ Module 1: Python FastAPI REST API Backend (All 8 API routers operational on port 8000).",
                        "✓ Module 2: Authentication & Role-Based Access Control (JWT tokens, bcrypt password hashing).",
                        "✓ Module 3: Dual Ingestion Engines (Java 21 Edge Gateway Daemon + Python IoT Simulator).",
                        "✓ Module 4: Interactive 3D WebGL Digital Twin (Real-time machine model with health-reactive animations).",
                        "✓ Module 5: Standalone Login Portal & Cyber Audio Synthesizer (Professional UI with synthesized sound effects)."
                    ]
                },
                {
                    "title": "Modules Currently Under Development",
                    "is_primary": False,
                    "bullets": [
                        "⚙ Module 6: Advanced Deep Learning LSTM Model for multi-step ahead RUL time-series forecasting (80% complete).",
                        "⚙ Module 7: Automated PDF Maintenance Report Generation for weekly plant audits (65% complete).",
                        "⚙ Module 8: Mobile Push Notification Gateway (Telegram/Twilio SMS integration for critical machine sirens) (50% complete)."
                    ]
                }
            ]
        },
        # SLIDE 15: Application Demonstrations
        {
            "category": "Chapter 6: Initial Implementation & Progress",
            "title": "Application Interface & Live Working Screens",
            "cards": [
                {
                    "title": "3D Digital Twin & Live Telemetry",
                    "is_primary": True,
                    "bullets": [
                        "Spatial Canvas: Real-time 3D rendered CNC machine reacting to sensor status (Green=Healthy, Amber=Warning, Red=Critical).",
                        "Telemetry Charts: Smooth Chart.js canvas visualizing vibration velocity RMS, motor temperature, and electrical current.",
                        "Cyber Audio Feedback: Interactive audio sound effects for machine selection, emergency sirens, and status changes.",
                        "Health Index Gauge: Color-coded radial progress bar showing machine condition from 100% (Optimal) to 0% (Fatal)."
                    ]
                },
                {
                    "title": "Enterprise Portal & API Capabilities",
                    "is_primary": False,
                    "bullets": [
                        "Secure Login Gateway: Glassmorphic auth portal with demo quick-fill credentials for examiners and evaluators.",
                        "Interactive Swagger UI: Full OpenAPI specification available at /api/v1/docs with live test sandbox.",
                        "One-Click Launcher: Complete platform launchable via 'start_all.bat' in 3 seconds on standard Windows.",
                        "Instant Free Public HTTPS Domain: Cloudflare tunnel script ('run_free_domain.bat') for remote live evaluation."
                    ]
                }
            ]
        },
        # SLIDE 16: Chapter 7 - Challenges & Future Plan
        {
            "category": "Chapter 7: Challenges & Future Work Plan",
            "title": "Technical Challenges & Future Timeline",
            "cards": [
                {
                    "title": "Technical Challenges & Solutions",
                    "is_primary": False,
                    "bullets": [
                        "Challenge 1 (Sensor Noise): Raw vibration data contained mechanical jitter. Solved using moving average exponential filtering.",
                        "Challenge 2 (Real-Time 3D Performance): High polygon models caused frame drops. Solved by optimizing Three.js meshes and lighting.",
                        "Challenge 3 (Dual Language Sync): Synchronizing Java edge gateway with Python FastAPI. Solved via standardized JSON REST schemas."
                    ]
                },
                {
                    "title": "Work Plan Before Final Presentation",
                    "is_primary": True,
                    "bullets": [
                        "Milestone 1 (Weeks 1-2): Complete LSTM model training and integrate with live prediction router.",
                        "Milestone 2 (Weeks 3-4): Implement PDF maintenance report generation and export.",
                        "Milestone 3 (Weeks 5-6): Complete extensive load testing (100+ concurrent simulated sensors).",
                        "Milestone 4 (Weeks 7-8): Finalize comprehensive documentation, user manual, and dissertation report."
                    ]
                }
            ]
        },
        # SLIDE 17: Conclusion
        {
            "category": "Conclusion",
            "title": "Conclusion & Project Summary",
            "cards": [
                {
                    "title": "Summary of Work Done",
                    "is_primary": True,
                    "bullets": [
                        "FactoryMind AI successfully demonstrates a working cyber-physical predictive maintenance platform for modern industry.",
                        "Achieved real-time telemetry streaming from Java edge gateway and Python simulator to a centralized AI analytics backend.",
                        "Delivered an engaging 3D Digital Twin web console providing instant spatial machine diagnostics and audio feedback.",
                        "Validates the feasibility of low-cost, open-standard AI condition monitoring for manufacturing enterprises."
                    ]
                },
                {
                    "title": "Academic Contribution & Value",
                    "is_primary": False,
                    "bullets": [
                        "Exemplifies practical application of MCA curriculum: distributed systems, cloud APIs, machine learning, and web graphics.",
                        "Provides a robust foundation for final semester implementation, testing, and research publication.",
                        "Acknowledged Guidance: Sincere gratitude to our project guide and Department of Computer Applications, TKIET, Warananagar.",
                        "Ready for Review-II committee evaluation and interactive live demonstration."
                    ]
                }
            ]
        },
        # SLIDE 18: References
        {
            "category": "References",
            "title": "References & Citations (IEEE Format)",
            "cards": [
                {
                    "title": "Research Literature & Datasets",
                    "is_primary": True,
                    "bullets": [
                        "[1] A. K. S. Jardine, D. Lin, and D. Banjevic, 'A review on machinery diagnostics and prognostics implementing condition-based maintenance,' Mechanical Systems and Signal Processing, vol. 20, no. 7, pp. 1483–1510, 2006.",
                        "[2] J. Lee, H. A. Kao, and S. Yang, 'Service innovation and smart analytics for Industry 4.0 and big data environment,' Procedia CIRP, vol. 16, pp. 3–8, 2014.",
                        "[3] A. Saxena and K. Goebel, 'Turbofan Engine Degradation Simulation Data Set,' NASA Ames Prognostics Data Repository, NASA Ames Research Center, Moffett Field, CA, 2008."
                    ]
                },
                {
                    "title": "Official Framework Documentation",
                    "is_primary": False,
                    "bullets": [
                        "[4] S. Ramírez, 'FastAPI Documentation: High-performance Python web framework,' tiangolo.com/fastapi, 2024.",
                        "[5] F. Pedregosa et al., 'Scikit-learn: Machine Learning in Python,' Journal of Machine Learning Research, vol. 12, pp. 2825–2830, 2011.",
                        "[6] Three.js Authors, 'Three.js: JavaScript 3D WebGL Library Documentation,' threejs.org, 2024.",
                        "[7] Oracle Corporation, 'Java Platform, Standard Edition Documentation - Release 21 LTS,' oracle.com/java, 2023."
                    ]
                }
            ]
        }
    ]

    total_slides = len(slides_content)

    for idx, sc in enumerate(slides_content, start=1):
        slide = prs.slides.add_slide(blank_layout)

        if sc.get("is_cover"):
            # Deep Navy Cover Slide Background
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = C_NAVY_DARK
            bg.line.fill.background()

            # Top Badge / Institution
            tb_inst = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.9))
            tf_inst = tb_inst.text_frame
            tf_inst.word_wrap = True
            p_inst = tf_inst.paragraphs[0]
            p_inst.text = sc["institution"]
            p_inst.font.size = Pt(13)
            p_inst.font.bold = True
            p_inst.font.color.rgb = RGBColor(147, 197, 253)

            # Main Title & Subtitle Box
            tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.733), Inches(2.2))
            tf_title = tb_title.text_frame
            tf_title.word_wrap = True
            
            p_title = tf_title.paragraphs[0]
            p_title.text = sc["title"]
            p_title.font.size = Pt(40)
            p_title.font.bold = True
            p_title.font.color.rgb = C_WHITE
            p_title.space_after = Pt(6)

            p_sub = tf_title.add_paragraph()
            p_sub.text = sc["subtitle"]
            p_sub.font.size = Pt(18)
            p_sub.font.color.rgb = RGBColor(226, 232, 240)
            p_sub.space_after = Pt(10)

            p_cat = tf_title.add_paragraph()
            p_cat.text = f"★  {sc['category']}"
            p_cat.font.size = Pt(13)
            p_cat.font.bold = True
            p_cat.font.color.rgb = RGBColor(56, 189, 248)

            # Members Card (Left)
            shape_mem = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.3), Inches(5.6), Inches(2.4))
            shape_mem.fill.solid()
            shape_mem.fill.fore_color.rgb = C_NAVY_CARD
            shape_mem.line.color.rgb = RGBColor(51, 65, 85)
            tf_mem = shape_mem.text_frame
            tf_mem.word_wrap = True
            tf_mem.margin_left = Inches(0.25)
            tf_mem.margin_top = Inches(0.2)

            p_mem_head = tf_mem.paragraphs[0]
            p_mem_head.text = "PRESENTED BY (SY MCA STUDENTS):"
            p_mem_head.font.size = Pt(12)
            p_mem_head.font.bold = True
            p_mem_head.font.color.rgb = RGBColor(148, 163, 184)
            p_mem_head.space_after = Pt(8)

            for m in sc["members"]:
                p = tf_mem.add_paragraph()
                p.text = m
                p.font.size = Pt(13)
                p.font.color.rgb = C_WHITE
                p.space_after = Pt(4)

            # Guide Card (Right)
            shape_g = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(4.3), Inches(5.7), Inches(2.4))
            shape_g.fill.solid()
            shape_g.fill.fore_color.rgb = C_NAVY_CARD
            shape_g.line.color.rgb = RGBColor(51, 65, 85)
            tf_g = shape_g.text_frame
            tf_g.word_wrap = True
            tf_g.margin_left = Inches(0.25)
            tf_g.margin_top = Inches(0.2)

            p_g_head = tf_g.paragraphs[0]
            p_g_head.text = "UNDER THE GUIDANCE OF:"
            p_g_head.font.size = Pt(12)
            p_g_head.font.bold = True
            p_g_head.font.color.rgb = RGBColor(148, 163, 184)
            p_g_head.space_after = Pt(8)

            for g_line in sc["guide"]:
                p = tf_g.add_paragraph()
                p.text = g_line
                p.font.size = Pt(13)
                p.font.color.rgb = C_WHITE
                p.space_after = Pt(4)

        else:
            # White Slide Background
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = C_WHITE
            bg.line.fill.background()

            # Add Standard Header & Footer
            add_header(slide, sc["category"], sc["title"], idx, total_slides)

            # Add Cards
            cards = sc.get("cards", [])
            if len(cards) == 2:
                # 2 Columns Layout
                col_w = 5.7
                col_h = 5.2
                add_card(slide, 0.8, 1.65, col_w, col_h, cards[0]["title"], cards[0]["bullets"], is_primary=cards[0].get("is_primary", True))
                add_card(slide, 6.8, 1.65, col_w, col_h, cards[1]["title"], cards[1]["bullets"], is_primary=cards[1].get("is_primary", False))
            elif len(cards) == 1:
                # Full Width
                add_card(slide, 0.8, 1.65, 11.733, 5.2, cards[0]["title"], cards[0]["bullets"], is_primary=cards[0].get("is_primary", True))

    prs.save(output_path)
    print(f"[SUCCESS] Native Microsoft PowerPoint Presentation saved to: {output_path}")

if __name__ == "__main__":
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_file = os.path.join(out_dir, "FactoryMind_AI_SY_MCA_Second_Presentation.pptx")
    create_deck(out_file)
