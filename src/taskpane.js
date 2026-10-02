let originalOoxml = null;
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
    setStatus("Applying FocusRead...");

    await Word.run(async context => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const selectedText = (selection.text || "").trim();
      if (!selectedText) {
        throw new Error("Select some text in Word first.");
      }

      originalOoxml = selection.getOoxml();
      await context.sync();

      const percentage = Number(document.getElementById("strength").value);

      const ranges = selection.split([" ", "\t", "\r", "\n"], true, true, true);
      ranges.load("items/text");
      await context.sync();

      const sample = ranges.items.slice(0, 5).map(r => JSON.stringify(r.text)).join(" | ");
      setStatus("Word found " + ranges.items.length + " text ranges: " + sample);

      let changed = 0;

      for (const range of ranges.items) {
        const word = (range.text || "").trim();

        if (!word || !/[A-Za-z0-9À-ÿ]/.test(word)) {
          continue;
        }

        const letters = word.match(/[A-Za-z0-9À-ÿ]/g);
        if (!letters || !letters.length) {
          continue;
        }

        const count = Math.max(1, Math.ceil(letters.length * percentage / 100));

        // Find the first count characters of the word within this word-range.
        const prefix = word.slice(0, count);

        const matches = range.search(prefix, {
          matchCase: true,
          matchWholeWord: false
        });

        matches.load("items");
        await context.sync();

        if (matches.items.length > 0) {
          // The range represents one word, so the first match is the
          // beginning of that word.
          matches.items[0].font.bold = true;
          changed++;
        }
      }

      await context.sync();

      if (!changed) {
        throw new Error("FocusRead could not find any word beginnings in the selected text.");
      }
    });

    setStatus("FocusRead applied — beginning of each word emphasised.");
  } catch (error) {
    console.error("FocusRead error:", error);
    setStatus(error && error.message ? error.message : "FocusRead could not format the selection.", true);
  } finally {
    busy = false;
  }
}

async function restoreOriginal() {
  if (!originalOoxml) {
    setStatus("There is no FocusRead formatting to restore yet.", true);
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
    setStatus("Original formatting restored.");
  } catch (error) {
    console.error("FocusRead restore error:", error);
    setStatus(error && error.message ? error.message : "Could not restore the original formatting.", true);
  } finally {
    busy = false;
  }
}
