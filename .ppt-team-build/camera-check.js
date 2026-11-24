
      const uploadModeButton = document.getElementById("upload-mode-button");
      const cameraModeButton = document.getElementById("camera-mode-button");
      const uploadMode = document.getElementById("upload-mode");
      const cameraMode = document.getElementById("camera-mode");
      const fileInput = document.querySelector('input[type="file"]');
      const filePrompt = document.getElementById("file-prompt");
      const fileDetails = document.getElementById("file-details");
      const startButton = document.getElementById("start-camera");
      const stopButton = document.getElementById("stop-camera");
      const captureButton = document.getElementById("capture-photo");
      const retakeButton = document.getElementById("retake-photo");
      const predictButton = document.getElementById("predict-photo");
      const video = document.getElementById("camera-feed");
      const capturedPhoto = document.getElementById("captured-photo");
      const canvas = document.getElementById("camera-canvas");
      const context = canvas.getContext("2d");
      const placeholder = document.getElementById("camera-placeholder");
      const status = document.getElementById("camera-status");
      const modelSelect = document.getElementById("model-select");
      const liveClass = document.getElementById("live-classification");
      const liveConfidence = document.getElementById("live-confidence");
      const livePredictions = document.getElementById("live-predictions");
      const unverifiedDialog = document.getElementById("unverified-dialog");
      const unverifiedMessage = document.getElementById("unverified-message");
      const closeUnverifiedDialog = document.getElementById("close-unverified-dialog");
      const installButton = document.getElementById("install-app");
      let cameraStream;
      let requestInFlight = false;
      let capturedImage;
      let installPrompt;

      function setMode(mode) {
        const cameraSelected = mode === "camera";
        uploadMode.hidden = cameraSelected;
        cameraMode.hidden = !cameraSelected;
        uploadModeButton.classList.toggle("is-active", !cameraSelected);
        cameraModeButton.classList.toggle("is-active", cameraSelected);
        uploadModeButton.setAttribute("aria-selected", String(!cameraSelected));
        cameraModeButton.setAttribute("aria-selected", String(cameraSelected));
        if (!cameraSelected && cameraStream) stopCamera();
      }

      function renderPredictions(predictions) {
        livePredictions.replaceChildren(...predictions.map((prediction) => {
          const row = document.createElement("div");
          row.className = "prediction";
          row.innerHTML = `
            <div class="prediction-label">
              <span>${prediction.label}</span>
              <span>${prediction.percent}%</span>
            </div>
            <div class="bar"><span style="width: ${prediction.percent}%"></span></div>
          `;
          return row;
        }));
      }

      function showUnverifiedDialog(topPrediction, thresholdPercent) {
        unverifiedMessage.textContent = `The best match was ${topPrediction.label} at ${topPrediction.percent}%, below the ${thresholdPercent}% minimum. Try a brighter, closer photo with one item in view.`;
        if (typeof unverifiedDialog.showModal === "function") {
          unverifiedDialog.showModal();
        } else {
          window.alert(unverifiedMessage.textContent);
        }
      }

      function capturePhoto() {
        if (!cameraStream || video.readyState < HTMLMediaElement.HAVE_CURRENT_DATA || !video.videoWidth || !video.videoHeight) {
          status.textContent = "Waiting for the camera picture. Keep the camera open for a moment, then try again.";
          return;
        }
        requestInFlight = true;
        const maxWidth = 640;
        const scale = Math.min(1, maxWidth / video.videoWidth);
        canvas.width = Math.max(1, Math.round(video.videoWidth * scale));
        canvas.height = Math.max(1, Math.round(video.videoHeight * scale));
        context.drawImage(video, 0, 0, canvas.width, canvas.height);
        capturedImage = canvas.toDataURL("image/jpeg", 0.82);
        capturedPhoto.src = capturedImage;
        capturedPhoto.hidden = false;
        video.hidden = true;
        captureButton.hidden = true;
        retakeButton.hidden = false;
        predictButton.hidden = false;
        liveClass.textContent = "Photo captured";
        liveConfidence.textContent = "Review the image, then select Predict image.";
        livePredictions.replaceChildren();
        status.textContent = "Photo captured. Review it before prediction.";
        requestInFlight = false;
      }

      function retakePhoto() {
        capturedImage = undefined;
        capturedPhoto.removeAttribute("src");
        capturedPhoto.hidden = true;
        video.hidden = false;
        captureButton.hidden = false;
        retakeButton.hidden = true;
        predictButton.hidden = true;
        liveClass.textContent = "Ready to scan";
        liveConfidence.textContent = "Capture an item when it is centered in the preview.";
        livePredictions.replaceChildren();
        status.textContent = "Camera is on. Center an item and capture it.";
      }

      async function predictPhoto() {
        if (!cameraStream || !capturedImage || requestInFlight) return;
        requestInFlight = true;
        predictButton.disabled = true;
        status.textContent = "Analyzing captured photo...";

        try {
          const response = await fetch("/api/predict-frame", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              model: modelSelect.value,
              image: capturedImage,
            }),
          });
          const data = await response.json();
          if (!response.ok) throw new Error(data.error || "Prediction failed.");

          liveClass.textContent = data.confident ? data.top_prediction.label : "Not confidently classified";
          liveConfidence.textContent = data.confident
            ? `${data.top_prediction.percent}% model score`
            : `Best match: ${data.top_prediction.label} at ${data.top_prediction.percent}% (below ${data.threshold_percent}% minimum)`;
          renderPredictions(data.predictions);
          status.textContent = "Prediction complete. Retake to scan another item.";
          if (!data.confident) showUnverifiedDialog(data.top_prediction, data.threshold_percent);
        } catch (error) {
          status.textContent = error.message;
        } finally {
          requestInFlight = false;
          predictButton.disabled = false;
        }
      }

      function waitForVideoFrame(timeoutMs = 8000) {
        if (video.readyState >= HTMLMediaElement.HAVE_CURRENT_DATA && video.videoWidth > 0 && video.videoHeight > 0) {
          return Promise.resolve();
        }

        return new Promise((resolve, reject) => {
          const timeout = window.setTimeout(() => {
            cleanup();
            reject(new Error("Camera opened but no picture arrived. Check that another app is not using the camera."));
          }, timeoutMs);
          const onReady = () => {
            if (video.videoWidth > 0 && video.videoHeight > 0) {
              cleanup();
              resolve();
            }
          };
          const cleanup = () => {
            window.clearTimeout(timeout);
            video.removeEventListener("loadeddata", onReady);
            video.removeEventListener("resize", onReady);
          };
          video.addEventListener("loadeddata", onReady);
          video.addEventListener("resize", onReady);
          onReady();
        });
      }

      async function startCamera() {
        if (!navigator.mediaDevices?.getUserMedia) {
          status.textContent = window.isSecureContext
            ? "This browser does not support camera access. Try uploading a photo instead."
            : "Camera access requires HTTPS or localhost. Open the secure site or upload a photo.";
          return;
        }

        try {
          status.textContent = "Requesting camera access...";
          const stream = await navigator.mediaDevices.getUserMedia({
            video: {
              facingMode: { ideal: "environment" },
              width: { ideal: 1280 },
              height: { ideal: 720 },
            },
            audio: false,
          });

          cameraStream = stream;
          video.srcObject = stream;
          video.hidden = false;
          await video.play();
          await waitForVideoFrame();
          placeholder.hidden = true;
          startButton.disabled = true;
          stopButton.disabled = false;
          captureButton.disabled = false;
          status.textContent = "Camera is on. The live picture is ready; center an item and capture it.";
          retakePhoto();
        } catch (error) {
          cameraStream?.getTracks().forEach((track) => track.stop());
          cameraStream = undefined;
          video.srcObject = null;
          placeholder.hidden = false;
          startButton.disabled = false;
          stopButton.disabled = true;
          captureButton.disabled = true;
          const messages = {
            NotAllowedError: "Camera permission is blocked. Allow camera access in browser settings, then try again.",
            NotFoundError: "No camera was found on this device. Try uploading a photo instead.",
            NotReadableError: "The camera is busy in another app. Close that app, then try again.",
            OverconstrainedError: "The requested camera mode is unavailable. Check your camera and try again.",
            SecurityError: "Camera access requires HTTPS or localhost. Open the secure site or upload a photo.",
          };
          status.textContent = messages[error.name] || error.message || "Could not start the camera. Check browser permission and try again.";
        }
      }

      function stopCamera() {
        cameraStream?.getTracks().forEach((track) => track.stop());
        cameraStream = undefined;
        video.srcObject = null;
        video.hidden = false;
        capturedPhoto.hidden = true;
        capturedPhoto.removeAttribute("src");
        capturedImage = undefined;
        placeholder.hidden = false;
        startButton.disabled = false;
        stopButton.disabled = true;
        captureButton.disabled = true;
        captureButton.hidden = false;
        retakeButton.hidden = true;
        predictButton.hidden = true;
        liveClass.textContent = "Ready to scan";
        liveConfidence.textContent = "Start the camera, capture an item, then predict.";
        livePredictions.replaceChildren();
        if (status.textContent !== "Camera is off.") status.textContent = "Camera stopped.";
      }

      fileInput.addEventListener("change", () => {
        const [file] = fileInput.files;
        if (!file) return;
        filePrompt.textContent = file.name;
        fileDetails.textContent = `${Math.ceil(file.size / 1024)} KB selected`;
      });
      uploadModeButton.addEventListener("click", () => setMode("upload"));
      cameraModeButton.addEventListener("click", () => setMode("camera"));
      startButton.addEventListener("click", startCamera);
      stopButton.addEventListener("click", stopCamera);
      captureButton.addEventListener("click", capturePhoto);
      retakeButton.addEventListener("click", retakePhoto);
      predictButton.addEventListener("click", predictPhoto);
      closeUnverifiedDialog.addEventListener("click", () => unverifiedDialog.close());
      unverifiedDialog.addEventListener("click", (event) => {
        if (event.target === unverifiedDialog) unverifiedDialog.close();
      });
      window.addEventListener("beforeunload", stopCamera);

      window.addEventListener("beforeinstallprompt", (event) => {
        event.preventDefault();
        installPrompt = event;
        installButton.hidden = false;
      });
      installButton.addEventListener("click", async () => {
        if (!installPrompt) return;
        installPrompt.prompt();
        await installPrompt.userChoice;
        installPrompt = undefined;
        installButton.hidden = true;
      });
      window.addEventListener("appinstalled", () => {
        installButton.hidden = true;
        status.textContent = "App installed. You can open it from your home screen.";
      });
      if ("serviceWorker" in navigator) {
        window.addEventListener("load", () => navigator.serviceWorker.register("/service-worker.js"));
      }
      
      showUnverifiedDialog({ label: "plastic", percent: 20 }, 40);
      
    