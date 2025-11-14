// Main documentation app functionality
document.addEventListener("DOMContentLoaded", function () {
    // Elements
    const searchBar = document.getElementById("search-bar");
    const topicList = document.getElementById("topic-list");
    const contentFrame = document.getElementById("content-frame");
    const toggleSidebarBtn = document.getElementById("toggle-sidebar");
    const sidebar = document.getElementById("sidebar");

    // Topic data structure
    const topicSections = [
        {
            header: "Getting Started",
            topics: [
                { name: "Introduction", file: "pages/introduction.html" },
                { name: "Installation", file: "pages/installation.html" },
            ],
        },
        {
            header: "Features",
            topics: [
                { name: "Basic Features", file: "pages/basic_features.html" },
                {
                    name: "Advanced Features",
                    file: "pages/advanced_features.html",
                },
            ],
        },
        {
            header: "Other",
            topics: [
                { name: "Troubleshooting", file: "pages/troubleshooting.html" },
                { name: "Architecture", file: "pages/architecture.html" },
                { name: "Contributions", file: "pages/contributions.html" },
                { name: "FAQ", file: "pages/faq.html" },
            ],
        },
    ];

    // Initialize topics list
    function populateTopics() {
        topicList.innerHTML = "";

        topicSections.forEach((section) => {
            // Create section header
            const headerItem = document.createElement("div");
            headerItem.className = "topic-header";
            headerItem.textContent = section.header;
            headerItem.setAttribute("data-type", "header");
            topicList.appendChild(headerItem);

            // Create topics under this section
            section.topics.forEach((topic) => {
                const topicItem = document.createElement("div");
                topicItem.className = "topic-item";
                topicItem.setAttribute("data-file", topic.file);

                const icon = document.createElement("i");
                icon.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="16" height="16"><path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/></svg>`;
                topicItem.appendChild(icon);

                const text = document.createTextNode(topic.name);
                topicItem.appendChild(text);

                topicList.appendChild(topicItem);

                // Add click event
                topicItem.addEventListener("click", function () {
                    loadTopic(topic.file, topic.name);
                    // Visual indicator for selected topic
                    document.querySelectorAll(".topic-item").forEach((item) => {
                        item.classList.remove("selected");
                    });
                    topicItem.classList.add("selected");
                });
            });
        });
    }

    // Load a topic into the content frame
    function loadTopic(filename, topicName) {
        // Show loading animation
        contentFrame.src = "pages/loading.html";

        // Track loading start time for analytics
        const startTime = performance.now();

        // Load the actual content after a delay (simulating or allowing for loading animation)
        setTimeout(() => {
            // Set up load event handler before changing the src
            contentFrame.onload = function () {
                // Calculate load time for analytics
                const loadTime = (
                    (performance.now() - startTime) /
                    1000
                ).toFixed(2);
                console.log(
                    `Page "${topicName}" loaded in ${loadTime} seconds`
                );

                try {
                    // Try to add keyboard listeners to the iframe content document
                    setupIframeKeyboardListeners(contentFrame);
                } catch (e) {
                    console.error(
                        "Error setting up iframe keyboard listeners:",
                        e
                    );
                }

                // Update page title based on current topic
                document.title = `Documentation - ${topicName}`;
            };

            // Now change the src to load the actual content
            contentFrame.src = filename;
        }, 800); // Slightly longer delay to see the animation
    }

    // Set up keyboard listeners for the iframe
    function setupIframeKeyboardListeners(iframe) {
        // Make sure the iframe and its document are accessible
        if (!iframe.contentWindow || !iframe.contentDocument) {
            return;
        }

        // Add keyboard event listener to the iframe's document
        iframe.contentWindow.addEventListener(
            "keydown",
            function (event) {
                // Ctrl+S to toggle sidebar
                if (event.ctrlKey && event.key === "s") {
                    event.preventDefault();
                    toggleSidebar();

                    // Focus search bar if sidebar is open
                    if (!sidebar.classList.contains("collapsed")) {
                        searchBar.focus();
                    }
                    return false;
                }

                // ESC to close
                if (event.key === "Escape") {
                    window.close();
                    return false;
                }
            },
            true
        );
    }

    // Filter topics based on search text
    function filterTopics(searchText) {
        const lowerSearchText = searchText.toLowerCase();
        let topicItems = document.querySelectorAll(
            ".topic-item, .topic-header"
        );

        let foundAny = false;
        let lastHeader = null;

        topicItems.forEach((item) => {
            if (item.getAttribute("data-type") === "header") {
                // This is a header, hide it initially
                item.style.display = "none";
                lastHeader = item;
            } else {
                // This is a topic item
                const matches = item.textContent
                    .toLowerCase()
                    .includes(lowerSearchText);
                item.style.display = matches ? "block" : "none";

                if (matches && lastHeader) {
                    // Show the header if we found a matching topic under it
                    lastHeader.style.display = "block";
                    foundAny = true;
                }
            }
        });

        return foundAny;
    }

    // Toggle sidebar visibility
    function toggleSidebar() {
        sidebar.classList.toggle("collapsed");
        document
            .getElementById("content-container")
            .classList.toggle("full-width");
    }

    // Event listeners
    if (searchBar) {
        searchBar.addEventListener("input", function () {
            filterTopics(this.value);
        });
    }

    if (toggleSidebarBtn) {
        toggleSidebarBtn.addEventListener("click", toggleSidebar);
    }

    // Main document keyboard shortcuts
    window.addEventListener(
        "keydown",
        function (event) {
            // Ctrl+S to toggle sidebar
            if (event.ctrlKey && event.key === "s") {
                event.preventDefault();
                toggleSidebar();

                // Focus search bar if sidebar is open
                if (!sidebar.classList.contains("collapsed")) {
                    searchBar.focus();
                }
                return false;
            }

            // ESC to close
            if (event.key === "Escape") {
                window.close();
                return false;
            }
        },
        true
    );

    // Initialize
    populateTopics();

    // Load initial content
    const initialTopic = topicSections[0].topics[0]; // Introduction
    loadTopic(initialTopic.file, initialTopic.name);

    // Performance metrics
    const loadTime =
        (window.performance.timing.domContentLoadedEventEnd -
            window.performance.timing.navigationStart) /
        1000;
    console.log(
        `Documentation app launch time: ${loadTime.toFixed(4)} seconds`
    );
});
