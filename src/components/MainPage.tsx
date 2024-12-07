import React, { useState } from "react";
import { Box, Button } from "@mui/material";
import SearchIcon from "@mui/icons-material/Search";
import AddIcon from "@mui/icons-material/Add";
// import axios from "axios";
import Popup from "./Popup";

const MainPage: React.FC = () => {
  const [popupOpen, setPopupOpen] = useState(false);
  const [apiData, setApiData] = useState<any[]>([]);

  // Function to open the external dialog with search slides data
  const handleSearchSlides = async () => {
  Office.onReady((info) => {
    try {
      const dummyData = [
        { id: 1, url: "https://raw.githubusercontent.com/Miraj-Tariq/abonea-python-test-Miraj/refs/heads/master/ppt_addin_icons/test_slide.png", name: "Slide 1" },
        { id: 2, url: "https://raw.githubusercontent.com/Miraj-Tariq/abonea-python-test-Miraj/refs/heads/master/ppt_addin_icons/test_slide.png", name: "Slide 2" },
        { id: 3, url: "https://raw.githubusercontent.com/Miraj-Tariq/abonea-python-test-Miraj/refs/heads/master/ppt_addin_icons/test_slide.png", name: "Slide 3" },
        { id: 4, url: "https://raw.githubusercontent.com/Miraj-Tariq/abonea-python-test-Miraj/refs/heads/master/ppt_addin_icons/test_slide.png", name: "Slide 4" },
        { id: 5, url: "https://raw.githubusercontent.com/Miraj-Tariq/abonea-python-test-Miraj/refs/heads/master/ppt_addin_icons/test_slide.png", name: "Slide 5" },
        { id: 6, url: "https://raw.githubusercontent.com/Miraj-Tariq/abonea-python-test-Miraj/refs/heads/master/ppt_addin_icons/test_slide.png", name: "Slide 6" },
      ];

      const dialogUrl = `https://localhost:3000/dialog.html?data=${encodeURIComponent(
        JSON.stringify(dummyData)
      )}`;

      Office.context.ui.displayDialogAsync(dialogUrl, { width: 60, height: 60 }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Dialog opened successfully");
        } else {
          console.error("Failed to open dialog:", result.error.message);
        }
      });
    } catch (error) {
      console.error("Error fetching search slides data", error);
    }
  });
};

  const handleCreateSlides = () => {
    console.log("Redirect to Create Slides Screen"); // Replace with your navigation logic
  };

  return (
    <Box display="flex" flexDirection="column" alignItems="center" padding="20px">
      <Button
        variant="outlined"
        startIcon={<SearchIcon />}
        fullWidth
        style={{ marginBottom: "10px" }}
        onClick={handleSearchSlides}
      >
        Search Slides
      </Button>
      <Button
        variant="outlined"
        startIcon={<AddIcon />}
        fullWidth
        onClick={handleCreateSlides}
      >
        Create Slides
      </Button>
      <Popup open={popupOpen} onClose={() => setPopupOpen(false)} data={apiData} />
    </Box>
  );
};

export default MainPage;
