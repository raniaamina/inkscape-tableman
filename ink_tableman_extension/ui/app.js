document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('extension-form');
    const submitBtn = document.getElementById('btn-submit');
    const cancelBtn = document.getElementById('btn-cancel');
    const statusContainer = document.getElementById('status-container');
    const progressBar = document.getElementById('progress-bar');
    const statusMessage = document.getElementById('status-message');

    // Dynamic Form Elements
    const shapeSelect = document.getElementById('shape_type');
    const textInputs = document.getElementById('text-inputs');
    const dimensionInputs = document.getElementById('dimension-inputs');

    // Handle view toggle based on shape selection
    shapeSelect.addEventListener('change', (e) => {
        const val = e.target.value;
        if (val === 'text') {
            textInputs.classList.remove('hidden');
            dimensionInputs.classList.add('hidden');
        } else {
            textInputs.classList.add('hidden');
            dimensionInputs.classList.remove('hidden');
        }
    });

    let pollInterval = null;

    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        // Collect data
        const formData = new FormData(form);
        const data = Object.fromEntries(formData.entries());

        // Update UI state
        submitBtn.disabled = true;
        statusContainer.classList.remove('hidden');
        updateProgress(5, "Submitting...", "normal");

        try {
            // Send trigger to Python backend
            const response = await fetch('/submit', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });

            if (response.ok) {
                // Begin polling status
                pollStatus();
            } else {
                updateProgress(0, "Failed to submit request.", "error");
                submitBtn.disabled = false;
            }
        } catch (err) {
            updateProgress(0, `Error: ${err.message}`, "error");
            submitBtn.disabled = false;
        }
    });

    cancelBtn.addEventListener('click', async () => {
        try {
            await fetch('/close', { method: 'POST' });
        } catch (err) {
            console.error(err);
        }
    });

    function updateProgress(percent, message, state = "normal") {
        progressBar.style.width = `${percent}%`;
        statusMessage.textContent = message;

        if (state === "error") {
            progressBar.classList.add('error');
        } else {
            progressBar.classList.remove('error');
        }
    }

    async function pollStatus() {
        if (pollInterval) clearInterval(pollInterval);

        pollInterval = setInterval(async () => {
            try {
                const response = await fetch('/status');
                const state = await response.json();

                updateProgress(state.progress, state.message, state.status === 'error' ? 'error' : 'normal');

                if (state.status === 'completed' || state.status === 'error') {
                    clearInterval(pollInterval);
                    submitBtn.disabled = false;
                }
            } catch (err) {
                console.error("Polling error:", err);
                clearInterval(pollInterval);
                submitBtn.disabled = false;
            }
        }, 500);
    }
});
