# Security Specification for Raport PAUD

## 1. Data Invariants
- A profile (User) cannot be modified unless the ID matches the owner or admin.
- Only Admins can modify configuration like `Kelas`, `TujuanPembelajaran`, `DataSekolah`.
- Only Gurus (or Admin) can modify academic records (`Siswa`, `NilaiIntrakurikuler`, `NilaiKokurikuler`, `CatatanAnak`, `Kehadiran`).

## 2. Dirty Dozen Payloads (Security Testing Hypotheses)
1. User profile creation with wrong UID.
2. User profile creation with ghost field `idRole`.
3. Guru updating a Kelas name.
4. Guru creating a TujuanPembelajaran.
5. User listing Siswa without authentication.
6. Admin deleting a Siswa field (if required).
7. Injecting 2000-character string into `Siswa.id`.
8. Updating `NilaiIntrakurikuler` with invalid score string.
9. Updating `Kehadiran` with negative number.
10. Spoofing admin by setting email manually in token (impossible, but tested).
11. Reading PII from another user profile (if required).
12. Creating a record with an invalid Class ID.

## 3. Security Audit Table

| Collection | Spoofing | State Shortcutting | Poisoning | Logic Leak |
| :--- | :--- | :--- | :--- | :--- |
| Users | Prevented | N/A | Prevented | Fixed |
| Kelas | Prevented | N/A | Prevented | Fixed |
| Siswa | Prevented | N/A | Prevented | Fixed |
| ... | ... | ... | ... | ... |

*Plan to Implement*:
1. Define validation helpers (`isValidSiswa`, etc.) based on `firebase-blueprint.json`.
2. Update rules to use these helpers in `allow write` blocks.
3. Clean up existing rules to fully comply with "Eight Pillars" of secure rules.
