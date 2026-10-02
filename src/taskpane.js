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

  document.getElementById("apply").addEventListener("click", testWordFormatting);
  document.getElementById("restore").addEventListener("click", restoreOriginal);
});

function setStatus(message, error = false) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.style.color = error ? "#a33" : "#66717d";
}

async function testWordFormatting() {
  if (busy) return;
  busy = true;

  try {
    setStatus("Testing Word formatting...");

    await Word.run(async context => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!(selection.text || "").trim()) {
        throw new Error("Select some text in Word first.");
      }

      originalOoxml = selection.getOoxml();
      await context.sync();

      selection.font.bold = true;
      await context.sync();
    });

    setStatus("Word connection works — selected text is now bold.");
  } catch (error) {
    console.error("FocusRead test error:", error);
    setStatus(error && error.message ? error.message : "Word formatting test failed.", true);
  } finally {
    busy = false;
  }
}

async function restoreOriginal() {
  if (!originalOoxml) {
    setStatus("There is no test formatting to restore yet.", true);
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
