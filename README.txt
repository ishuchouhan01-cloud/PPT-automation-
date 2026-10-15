EDOS WEEKLY REVIEW CLOUD
========================

WHAT IT DOES
Turns an ASTA Powerproject schedule export into an editable TotalEnvironment Weekly Review
PowerPoint deck. Every slide is real PowerPoint content, so you can edit it afterwards.


TWO PROJECT TYPES
You choose one at the top of the page.

  Launch & Design    Plotted and apartment launches. Includes Approvals Status,
                     Land Acquisition and Master Plan.

  Design & Build     Individual residences and campuses. No Approvals Status slide;
                     adds Schedule - Construction and Project Overview - Construction
                     where the project has them.


WHAT YOU NEED INSTALLED
  Node.js LTS (18 or newer).  Free, from https://nodejs.org
  Nothing else. See REQUIREMENTS.txt.


HOW TO START
  1. Extract this ZIP to any folder you can write to.
  2. Double-click START.bat.
  3. Your browser opens at http://127.0.0.1:3000
  4. Leave the black window open while you work. Closing it stops the application.


LETTING SOMEONE ELSE TRY IT
  Double-click SHARE.bat instead. It does everything START.bat does and also puts a
  public web address in front of it, so a tester anywhere can open the application
  running on this laptop. They need nothing installed.

  Set a password in sharing.json first — SHARE.bat will not start without one.
  The address is temporary and dies when you close the window.

  Full instructions, and what to know before you send the link: SHARING.txt


WHERE YOU PUT YOUR FILES
Upload them on the page itself — drag and drop, or click a box. Nothing is copied into
this folder by hand, and nothing is sent over the internet: it all stays on your machine.

  Launch & Design   ASTA Excel export ...... required
                    Summary Dashboard ...... for Red Flag, Key Focus and the Tracker
                    Previous Weekly PPT .... for the cover and carried-over slides
                    Land Acquisition ....... only if that slide changed this week

  Design & Build    Excel export view ...... required, and the only schedule file. Its
                                             Design and Construction branches are read
                                             from inside it — no separate construction upload
                    Summary Dashboard ...... for Red Flag, Key Focus and both timelines
                    Previous Weekly PPT .... for the cover and Site Progress pages


WHERE THE FINISHED DECK GOES
Click "Download PPTX" when generation finishes — it saves to your normal Downloads folder.
A copy is also kept in this package under  storage\generated\  next to its validation report,
so you can find an earlier run without regenerating it.


IF SOMETHING GOES WRONG
See TROUBLESHOOTING.txt. The commonest cause by far is Node.js not being installed.


A NOTE ON FONTS
The deck is designed in Impact, Helvetica Neue and Meta-Normal. The fonts\ folder holds the
approved files — install them on any machine that will open or present the deck, or PowerPoint
will substitute something else and the layout will shift. Right-click each file, then Install.
