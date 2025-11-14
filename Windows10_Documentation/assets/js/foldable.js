document.addEventListener("DOMContentLoaded", function () {
    var headers = document.querySelectorAll("h1, h2, h3, h4, h5, h6");

    headers.forEach(function (header) {
        // Find all content elements that belong to this header
        var contentElements = [];
        var currentElement = header.nextElementSibling;
        var buttonElements = [];

        while (currentElement && !currentElement.tagName.match(/^H[1-6]$/)) {
            // Specifically identify button elements to handle separately
            if (currentElement.classList.contains("btn")) {
                buttonElements.push(currentElement);
            } else {
                contentElements.push(currentElement);
            }
            currentElement = currentElement.nextElementSibling;
        }

        if (contentElements.length > 0) {
            // Create a container div for all content elements
            var container = document.createElement("div");
            container.className = "header-content";

            // Add arrow element
            var arrow = document.createElement("span");
            arrow.className = "header-arrow";
            arrow.innerHTML = "▼";

            // Insert the container after the header
            header.parentNode.insertBefore(
                container,
                header.nextElementSibling
            );

            // Move all content elements into the container
            contentElements.forEach(function (element) {
                container.appendChild(element);
            });

            // Place button elements after the container
            buttonElements.forEach(function (button) {
                header.parentNode.insertBefore(button, container.nextSibling);
            });

            header.insertBefore(arrow, header.firstChild);

            // Set initial state based on header level
            if (header.tagName === "H3") {
                container.classList.add("collapsed");
                arrow.classList.add("collapsed");
            }

            // Add click event listener to toggle visibility
            header.addEventListener("click", function (event) {
                // Prevent default behavior if header is also a link
                if (event.target.tagName === "A") {
                    return;
                }

                event.preventDefault();

                if (container.classList.contains("collapsed")) {
                    container.classList.remove("collapsed");
                    arrow.classList.remove("collapsed");
                } else {
                    container.classList.add("collapsed");
                    arrow.classList.add("collapsed");
                }
            });
        }
    });
});
