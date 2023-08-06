import React from "react";
import { Link } from "react-router-dom";
import "./nav.scss";

export default function Nav() {
  return (
    <nav className="navbar navbar-expand-lg bg-dark">
      <div className="container-fluid">
        <Link to="/" className="text-light h5">
          الاختبارات
        </Link>
        <button
          className="navbar-toggler"
          type="button"
          data-bs-toggle="collapse"
          data-bs-target="#navbarNav"
        >
          <i class="bi bi-list text-light" />
        </button>
        <div className="collapse navbar-collapse" id="navbarNav">
          <ul className="navbar-nav">
            <li className="nav-item">
              <Link to="/" className="nav-link text-light">
                إنشاء إستمارات
              </Link>
              <Link to="/make-certificate" className="nav-link text-light">
                إصدار الشهادات
              </Link>
            </li>
          </ul>
        </div>
        <div className="collapse navbar-collapse" id="navbarNavAltMarkup">
          <div className="navbar-nav">
            <Link to="/" className="nav-link text-light">
              إنشاء إستمارات
            </Link>
            <Link to="/make-certificate" className="nav-link text-light">
              إصدار الشهادات
            </Link>
          </div>
        </div>
      </div>
    </nav>
  );
}
