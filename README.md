# FactoryMind AI — Industrial Predictive Manufacturing Platform

> **Academic Project:** SY MCA Major Project — Preliminary Project Report (Review-II)  
> **Institution:** Department of Computer Applications, Tatyasaheb Kore Institute of Engineering & Technology (TKIET), Warananagar  
> **Academic Year:** 2026–2027  
> **Standards Compliance:** ISO 13379-1 Condition Monitoring and Diagnostics of Machines

---

## ⚡ Instant 1-Click Launch (Zero Configuration Needed!)

Simply **double-click** the file:
```cmd
1_CLICK_RUN.bat
```

### What happens when you double-click:
1. It immediately launches the **FactoryMind AI Web Platform** on `http://127.0.0.1:8000/login`.
2. Automatically opens your default web browser (Google Chrome, Edge, or Brave).
3. **No Python, Java, or external software is needed!** It runs seamlessly on Windows out-of-the-box.

---

## 🔑 Demo Access Profiles

On the login portal, click any of the **Quick Fill Profile buttons** or the **RFID Tap** button for instant access:

| Role Profile | Name | Email | Password | Access Privileges |
| :--- | :--- | :--- | :--- | :--- |
| **Lead Reliability Engineer** | Rahul Sharma | `engineer@factorymind.io` | `engineer123` | 3D Twin Hotspots, FFT Spectrogram, Fault Injection |
| **Plant Operations Manager** | Priya Nair | `manager@factorymind.io` | `manager123` | Plant Health KPIs, MTBF/MTTR, ISO 13379 Shift Reports |
| **Root Administrator** | Vikram Patel | `admin@factorymind.io` | `admin123` | Asset Provisioning, Sensor Thresholds, User Roles |

---

## 🌟 Interactive Features in the Platform

1. **2.5D Plant Floor:** Live status of all 6 machines across Bays A through E with animated conveyor flows.
2. **3D WebGL Digital Twin:** Interactive 3D Lathe & Milling Spindle (rotate, zoom, inspect component hotspots for bearings, motor casing, and drive coupling).
3. **60 FPS Live FFT Spectrogram:** Real-time frequency harmonic analysis showing 1X, 2X, 3X, and BPFO bearing defect spikes.
4. **Interactive Fault Injector:** Click **"Fault Injector"** in the top bar to simulate real-world bearing wear, thermal overload, or electrical imbalance, triggering warning sirens and RUL drops.
5. **Conversational AI Copilot:** Ask questions about machine health, why an asset is degrading, or request maintenance dispatch protocols.
6. **Maintenance Work Orders:** Dispatch, track, and complete work orders.
7. **ISO 13379-1 Shift Audit Report:** Printable shift audit document ready for review evaluation (`Ctrl+P` to save as PDF).

---

## 📁 Project Directory Structure

```
FactoryMind_AI/
├── 1_CLICK_RUN.bat                                           # 1-Click Instant Launcher (Zero Install!)
├── run_free_domain.bat                                       # Free Public Cloudflare HTTPS Tunnel
├── run_java_gateway.bat                                      # Java 21 PLC Edge Gateway
├── generate_presentation.py                                  # PowerPoint (.pptx) Deck Generator
├── presentation.html                                         # 16:9 Widescreen Browser Presentation Deck
├── README.md                                                 # Project Manual & Guide
├── java-edge-gateway/
│   └── src/com/factorymind/gateway/IndustrialEdgeGateway.java # Java Edge Daemon
└── backend/
    ├── server.ps1                                            # Built-in Windows Web Server
    ├── requirements.txt                                      # Python Dependencies
    ├── seed_data.py                                          # Database Seeder
    ├── static/
    │   ├── login.html                                        # Standalone Cyber Security Login Portal
    │   ├── index.html                                        # 3D Digital Twin & Operations Console
    │   ├── style.css                                         # Cyber-Industrial CSS
    │   └── app.js                                            # Three.js 3D & Web Audio Engine
    ├── simulator/
    │   └── iot_simulator.py                                  # Python Sensor Telemetry Streamer
    └── app/
        ├── main.py                                           # FastAPI REST Core
        ├── core/                                             # Database & Config
        ├── models/                                           # SQLAlchemy ORM Models
        ├── schemas/                                          # Pydantic Schemas
        ├── ml/                                               # Machine Learning Diagnostic Engine
        └── api/                                              # 8 REST API Routers
```
