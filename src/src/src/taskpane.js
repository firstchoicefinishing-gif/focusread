```javascript
let originalOoxml = null;
let busy = false;

Office.onReady(() => {
  const strength = document.getElementById("strength");
  const strengthValue = document.getElementById("strengthValue");

  strength.addEventListener("input", () => {
    strengthValue.textContent = strength.value + "%";

    document
      .querySelectorAll("[data-strength]")
      .forEach(button => button.classList.remove("active"));
  });

  document
    .querySelectorAll("[data-strength]")
    .forEach(button => {
      button.addEventListener("click", () => {
        strength.value = button.dataset.strength;
        strengthValue.textContent = strength.value + "%";

        document
          .querySelectorAll("[data-strength]")
          .forEach(item => item.classList.remove("active"));

        button.classList.add("active");
      });
    });

  document
    .getElementById("apply")
    .addEventListener("click", applyFocusRead);

  document
    .getElementById("restore")
    .addEventListener("click", restoreOriginal);
});

function setStatus(message, error = false) {
  const status = document.getElementById("status");

  status.textContent = message;
  status.style.color = error ? "#a33" : "#66717d";
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

      if (!selection.text || !selection.text.trim()) {
        throw new Error("Select some text in Word first.");
      }

      /*
       * Save the original formatting so the user can restore it.
       */
      originalOoxml = await selection.getOoxml();

      const percentage = Number(
        document.getElementById("strength").value
      );

      /*
       * Split the selected text into word-like ranges.
       */
      const ranges = selection.getTextRanges(
        [
          " ",
          "\t",
          "\r",
          "\n",
          ".",
          ",",
          ";",
          ":",
          "!",
          "?",
          "(",
          ")",
          "[",
          "]",
          "{",
          "}",
          "/",
          "\\",
          "-",
          "—",
          "–",
          "\"",
          "'"
        ],
        true
      );

      ranges.load("items/text");

      await context.sync();

      for (const range of ranges.items) {

        const text = range.text || "";
        const word = text.trim();

        if (!word) continue;

        /*
         * Calculate how much of the word should be emphasised.
         */
        const characters = Math.max(
          1,
          Math.ceil(word.length * percentage / 100)
        );

        const prefix = word.substring(0, characters);

        /*
         * Find the beginning of the word within the Word range.
         */
        const matches = range.search(prefix, {
          matchCase: true,
          matchWholeWord: false
        });

        matches.load("items");

        await context.sync();

        for (const match of matches.items) {
          match.font.bold = true;
        }
      }

      await context.sync();
    });

    setStatus(
      "FocusRead applied. Use Restore original when finished."
    );

  } catch (error) {

    setStatus(
      error.message || "Something went wrong.",
      true
    );

  } finally {
    busy = false;
  }
}

async function restoreOriginal() {

  if (!originalOoxml) {
    setStatus(
      "There is no saved selection to restore yet.",
      true
    );
    return;
  }

  if (busy) return;

  busy = true;

  try {

    setStatus("Restoring original formatting...");

    await Word.run(async context => {

      const selection = context.document.getSelection();

      selection.insertOoxml(
        originalOoxml,
        Word.InsertLocation.replace
      );

      await context.sync();
    });

    originalOoxml = null;

    setStatus("Original formatting restored.");

  } catch (error) {

    setStatus(
      error.message || "Could not restore the original formatting.",
      true
    );

  } finally {
    busy = false;
  }
}
```
