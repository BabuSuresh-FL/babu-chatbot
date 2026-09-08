# 🎈 Blank app template

A simple Streamlit app template for you to modify!

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://blank-app-template.streamlit.app/)

### How to run it on your own machine

1. Install the requirements

   ```
   $ pip install -r requirements.txt
   ```

2. Run the app

   ```
   $ streamlit run streamlit_app.py
   ```
======
Version-1
Mainframe Storage Engineering
Team Overview & Process Documentation
Field	Detail
Document Purpose	Introductory discussion capturing the Mainframe Storage Engineering team's scope, responsibilities, and infrastructure footprint
Source	Recorded conversation between three participants (interviewer and two Mainframe Storage Engineering team members)
Team Name	Mainframe Storage Engineering (also referred to informally as "Mainframe Storage")
1. Purpose
This document summarizes an introductory conversation intended to capture the scope, responsibilities, and supporting infrastructure of the Mainframe Storage Engineering team. It is organized by topic rather than by speaker, for use as reference documentation.
2. Team Identification
Official Team Name: Mainframe Storage Engineering
One team member noted that the team is generally referred to simply as "Mainframe Storage," with "Engineering" sometimes appended informally.
Team Tenure (participants): One team member has been with the team for approximately 1.5 years; the other has approximately 5 years of tenure and holds significant institutional/tribal knowledge of team structure and naming conventions.
3. Disk Storage — Architecture
•	Storage is delivered through multiple subsystems; even a single physical frame is logically composed of multiple frames.
•	Each subsystem provides petabytes to terabytes of storage.
•	Physical solid-state/flash drives within the frame are logically presented to the mainframe as Count-Key Data (CKD) volumes — emulating legacy spinning-platter devices (e.g., 15 platters forming one cylinder of 15 tracks).
•	All disk storage infrastructure is on premises.
4. Disk Storage — Installation & Setup Process
The setup process involves several teams working together:
•	A dedicated hardware team receives equipment and place it on the data center floor.
•	The hardware team and networking team coordinate the initial spin-up of devices (channels, ports, physical connections).
•	The Mainframe Storage Engineering team drives much of the overall process once hardware is in place.
•	Status is documented via recap emails and maintained on the team's SharePoint site, which serves as the central reference for infrastructure setup.
5. Ongoing Responsibilities — Disk
The team does not physically touch hardware. All storage-related work is performed logically/remotely:
•	Physical installation, cabling, and hands-on hardware work is performed exclusively by data center floor staff.
•	The hardware team manages requests and coordination but does not touch machines directly.
•	Once devices are spun up on the network, the team performs storage configuration, including disk initialization.
•	Devices are organized into Logical Storage Groups and made available via the team's Storage Management Software (SMS).
•	The team ensures storage is backed up and available for migration and space management.
•	The team resolves file allocation issues and related storage problems.
•	The team manages disaster recovery activities, including "gold copy" backups (referred to as LCP and SDC) used for vaulting-type purposes.
6. Tape / Virtual Tape Storage (VTS)
Tape follows a very similar process to disk, with a few distinctions:
•	Physical tape hardware is installed by data center staff; the team never physically touches it.
•	The team handles logical setup, confirms drives are present and operating correctly, and integrates the hardware into the storage management software environment so tapes can be accessed and mounted.
•	Tapes are virtual, not physical — they are backed by underlying disk storage.
•	The core difference from disk is in how the subsystem communicates with the mainframe; the overall process is otherwise very similar.
7. Geographic Footprint & Disaster Recovery
Coverage: United States only — no mainframe infrastructure outside the U.S.
Data Center Locations: Phoenix, Arizona and Dallas, Texas
High Availability / DR Model: Phoenix operates as the primary site; Dallas serves as the failover site in a high-availability configuration.
8. Systems & Environments
•	Multiple physical mainframe boxes exist across locations (referenced individually, e.g., M12, M13).
•	Each physical mainframe is subdivided into Logical Partitions (LPARs) logical, independent operating system instances running under a single physical CPU, conceptually like virtual machines (VMs) in a Windows or Linux environment.
•	A single physical mainframe box may run four or five separate logical systems.
•	Storage cabinets/frames are physically separate from — but connected to — the mainframe CPU cabinets.
•	Multiple operating environments exist per system, including Test, Dev, NFTS, and Production.
•	Across environments, there may be on the order of 1,000 file system iterations, with tens to hundreds of thousands of jobs running daily on the mainframe.
•	Detailed LPAR counts and configuration are owned by the Operating System (OS) team; the Mainframe Storage Engineering team's estimate is close to 100 LPARs managed across the physical mainframes.
9. Infrastructure Inventory (Team-Reported Estimates)
Item	Count / Detail
Linux boxes — Phoenix	2
Mainframe (M-series) boxes — Phoenix	5
Linux boxes — Texas	2
Mainframe boxes — Texas	4
Total physical boxes (Phoenix + Texas)	13
LPARs (approximate, team estimate)	Close to 100 — authoritative count owned by the OS team
Disk storage boxes (primary focus of this team)	29
VTS (Virtual Tape System) boxes (primary focus of this team)	8
Note: The team clarified that precise LPAR counts should be confirmed with the Operating System (OS) team, as LPAR management falls under that team's ownership. The disk and VTS box counts (29 and 8, respectively) represent the infrastructure the Mainframe Storage Engineering team is most directly responsible for.
10. Ongoing Operational Activity
Storage is not static once installed — the team described continuous lifecycle activity, including:
•	Regular migration of data into and out of storage.
•	Management across multiple operating system environments (Test, Dev, NFTS, Prod), each with numerous file system variations.
•	Movement of data between live DASD, near-live DASD, and migrated/cold storage to optimize capacity.
•	Daily management of gold copy backups.
•	Certificate management for various storage subsystems.
•	Coordination across multiple subsystems tied to different machines, LPARs, and physical mainframe boxes — including the VTS environment, which is logically composed of six separate clusters.
11. Summary
The Mainframe Storage Engineering team is responsible for the logical configuration, provisioning, backup, disaster recovery, and lifecycle management of mainframe disk and virtual tape storage across two U.S. data centers (Phoenix, primary; Dallas, failover). The team does not perform physical hardware installation — this is handled by data center floor staff and coordinated with the hardware and networking teams. The team's core infrastructure footprint includes 29 disk storage boxes and 8 VTS boxes, supporting mainframe environments spread across 13 physical mainframe/Linux boxes and roughly 100 LPARs.

Version-2

Mainframe Storage Architecture & Team Responsibilities
1. System & Architecture Overview
•	Disk Storage Architecture: Physical disk storage consists of multiple subsystems utilizing solid-state/flash drives. Logically, these drives are configured to emulate traditional Count-Key Data (CKD) volumes.

•	Virtual Tape Systems (VTS): Tape storage functions similarly to disk storage, utilizing virtual tapes backed directly by disk storage.

•	LPAR vs. Storage: An LPAR (Logical Partition) is a virtualized operating system instance running on the mainframe CPU hardware. Mainframe CPU cabinets house the CPUs and LPARs, while storage cabinets (DASD and VTS) are physically separate entities connected via specialized hardware interfaces.

•	Deployment Model: All storage infrastructure is hosted strictly on-premises across US-based data centers.

2. Infrastructure & Physical Asset Inventory
Primary Data Centers
•	Phoenix, AZ: Primary Data Center Location.

•	Texas (Dallas Area): Disaster Recovery (DR) and High-Availability Location.

Hardware Footprint
Category	Component / Asset Type	Count	Location / Details
CPU / Server Boxes	Mainframe Physical Boxes	9 Boxes	5 in Phoenix, 4 in Texas
	Linux Physical Boxes	4 Boxes	2 in Phoenix, 2 in Texas
Logical Partitioning	Operating System LPARs	~100 LPARs	Spans Dev, Test, NFTS, and Prod environments
Managed Storage Assets	Disk Storage Subsystems (DASD)	29 Cabinets	Physical storage frames managed by the storage team
	Virtual Tape Subsystems (VTS)	8 Cabinets	Physical virtual tape frames managed by the storage team
3. Team Division of Responsibilities
Mainframe Storage Engineering Team Scope
•	Logical Initialization & Setup: Initializing new disks, creating Logical Storage Groups (LSGs), and exposing storage through Storage Management Software (SMS).

•	Storage Operations & Life Cycle: Managing daily storage allocations, data movement across live and near-live storage, certificate management across subsystems, and lifecycle migrations.

•	Disaster Recovery & Backup Management: Managing logical copy backups (LCP), Gold Copy backups, and Synchronous Data Center Replication (SDC) for vaulting and recovery purposes.

•	Logical-Only Operations: The storage team works purely at the logical/virtual level and does not physically access or manage hardware on the data center floor.

Supporting Engineering & Operations Teams
•	Data Center Operations: Physical racking, cabling, and hands-on maintenance of machine hardware on the floor.

•	Hardware & Networking Teams: Physical hardware procurement, request routing, initial device spin-up, and physical channel/port/network connections.

•	Operating System Team: Management of CPU hardware partitioning, LPAR allocation, and OS-level configurations.


======
