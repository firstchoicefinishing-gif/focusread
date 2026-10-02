let originalOoxml = null;
let savedSelection = false;
let busy = false;

Office.onReady(() => {
  const strength = document.getElementById("strength");
  const strengthValue = document.getElementById("strengthValue");

  strength.addEventListener("input", () => {
    strengthValue.textContent = strength.value + "%";
    document.querySelectorAll("[data-strength]").forEach(b => b.classList.remove("active"));
  });

  document.querySelectorAll("[data-strength]").forEach(button => {
    button.addEventListener("click", () => {
      strength.value = button.dataset.strength;
      strengthValue.textContent = strength.value + "%";
      document.querySelectorAll("[data-strength]").forEach(b => b.classList.remove("active"));
      button.classList.add("active");
    });
  });

  document.getElementById("apply").addEventListener("click", applyFocusRead);
  document.getElementById("restore").addEventListener("click", restoreOriginal);
});

function setStatus(message, error = false) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.style.color = error ? "#a33" : "#66717d";
}

async function applyFocusRead() {
  if (busy) return;
  busy = true;

  try {
    setStatus("Reading selection...");

    await Word.run(async context => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = (selection.text || "").trim();
      if (!text) throw new Error("Select some text in Word first.");

      originalOoxml = selection.getOoxml();
      await context.sync();

      const percentage = Number(document.getElementById("strength").value);
      const ranges = selection.getTextRanges([" ", "\t", "\r", "\n", ".", ",", ";", ":", "!", "?", "(", ")", "[", "]", "{", "}", "/", "\\", "-", "—", "–", """, "'"], true);
      ranges.load("items/text");
      await context.sync();

      let changed = 0;

      for (const range of ranges.items) {
        const word = (range.text || "").trim();
        if (!word) continue;

        const count = Math.max(1, Math.ceil(word.length * percentage / 100));
        const prefix = word.slice(0, count);

        const matches = range.search(prefix, {
          matchCase: true,
          matchWholeWord: false
        });
        matches.load("items");
        await context.sync();

        for (const match of matches.items) {
          match.font.bold = true;
          changed++;
        }
      }

      await context.sync();
      savedSelection = true;

      if (!changed) {
        throw new Error("Word found the selection, but could not format the word beginnings.");
      }
    });

    setStatus("FocusRead applied. Use Restore original when finished.");
  } catch (error) {
    console.error("FocusRead error:", error);
    setStatus(error && error.message ? error.message : "FocusRead could not format the selection.", true);
  } finally {
    busy = false;
  }
}

async function restoreOriginal() {
  if (!savedSelection || !originalOoxml) {
    setStatus("There is no FocusRead selection to restore yet.", true);
    return;
  }

  if (busy) return;
  busy = true;

  try {
    setStatus("Restoring original formatting...");

    await Word.run(async context => {
      const selection = context.document.getSelection();
      selection.insertOoxml(originalOoxml.value, Word.InsertLocation.replace);
      await context.sync();
    });

    originalOoxml = null;
    savedSelection = false;
    setStatus("Original formatting restored.");
  } catch (error) {
    console.error("FocusRead restore error:", error);
    setStatus(error && error.message ? error.message : "Could not restore the original formatting.", true);
  } finally {
    busy = false;
  }
}