import React from "react";
import { BrowserRouter, Routes, Route } from "react-router-dom";
import "./App.css";
import { PDFDownloadLink } from "@react-pdf/renderer";

// Components
import Nav from "./Components/Nav/Nav";
import Certeficate from "./Components/Certificate/Certeficate";

function App() {
  return (
    <div>
      <BrowserRouter>
        <Nav />
        <Routes>
          <Route index path="/" element={<>hi</>} />
          <Route
            path="/make-certificate"
            element={
              <>
                <Certeficate />
                <div>
                  <PDFDownloadLink
                    document={<Certeficate />}
                    fileName="Certeficate.pdf"
                  >
                    {({ blob, url, loading, error }) =>
                      loading ? "Loading document..." : "Download now!"
                    }
                  </PDFDownloadLink>
                </div>
              </>
            }
          />
        </Routes>
      </BrowserRouter>
    </div>
  );
}

export default App;
