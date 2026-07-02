import React, { useState, useEffect, useRef } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx-js-style';
import { Settings, RefreshCw, FilePlus, Database, Download, ExternalLink, AlertCircle, X, ChevronDown, Eye, Lock, User as UserIcon, LogIn, LogOut, Filter } from 'lucide-react';
import { EVN_HCMC_LOGO } from "./assets/logo";
import { loadProjectData, checkDuplicateDocument } from "./services/dataService";

const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/14BF0RUfBq-Arl6ngVvD44fQnNayBEC1Xtz-RFgzA4GI/export?format=csv&gid=0';
const PROJECTS_CSV_URL = 'https://docs.google.com/spreadsheets/d/14BF0RUfBq-Arl6ngVvD44fQnNayBEC1Xtz-RFgzA4GI/export?format=csv&gid=1152018861'; // Using gid for 'Thông tin theo MCT' if known, or gviz. Let's use gviz to be safe.
const PROJECTS_GVIZ_URL = 'https://docs.google.com/spreadsheets/d/14BF0RUfBq-Arl6ngVvD44fQnNayBEC1Xtz-RFgzA4GI/gviz/tq?tqx=out:csv&sheet=' + encodeURIComponent('Thông tin theo MCT');
const USERS_GVIZ_URL = 'https://docs.google.com/spreadsheets/d/14BF0RUfBq-Arl6ngVvD44fQnNayBEC1Xtz-RFgzA4GI/gviz/tq?tqx=out:csv&sheet=' + encodeURIComponent('user');

const PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(SHEET_CSV_URL)}`;
const PROJECTS_PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(PROJECTS_GVIZ_URL)}`;
const USERS_PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(USERS_GVIZ_URL)}`;

const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx1X9iZX0ZWin-gB2ORtL-IkaCp8Px4WKaz4vV2ISr7jyuT9Wviuct0fR6sIJAs051n/exec';

interface User {
  FullName: string;
  username: string;
  AllowedProjects: string;
}

export default function App() {
  const [currentUser, setCurrentUser] = useState<User | null>(() => {
    const savedUser = localStorage.getItem('currentUser');
    return savedUser ? JSON.parse(savedUser) : null;
  });
  const [users, setUsers] = useState<any[]>([]);
  const [loginForm, setLoginForm] = useState({ username: '', password: '' });
  const [loginError, setLoginError] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const [data, setData] = useState<string[][]>([]);
  const [projectInfo, setProjectInfo] = useState<Record<string, string>>({
    "Tên dự án/công trình": "",
    "Mã công trình": "",
    "Chủ Đầu Tư": "",
    "Địa điểm xây dựng": "",
    "Đơn vị TV Thiết Kế": "",
    "Đơn vị TV Giám sát": "",
    "Đơn vị thi công": ""
  });
  const [availableProjects, setAvailableProjects] = useState<Record<string, string>[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [scriptUrl, setScriptUrl] = useState(localStorage.getItem('appsScriptUrl') || DEFAULT_SCRIPT_URL);
  const [showSettings, setShowSettings] = useState(false);
  const [showImportModal, setShowImportModal] = useState(false);
  const [actionLoading, setActionLoading] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [projectSearchTerm, setProjectSearchTerm] = useState('');
  const [lastUpdated, setLastUpdated] = useState<string>('');
  const [collapsedSections, setCollapsedSections] = useState<Set<string>>(new Set());
  const [selectedSection, setSelectedSection] = useState<string>('all');
  const [showMobileFilters, setShowMobileFilters] = useState(false);

  const [templateCache, setTemplateCache] = useState<string[][] | null>(null);

  const [showDropdown, setShowDropdown] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  const [showImportProjectDropdown, setShowImportProjectDropdown] = useState(false);
  const [importProjectSearchTerm, setImportProjectSearchTerm] = useState('');
  const importProjectDropdownRef = useRef<HTMLDivElement>(null);
  
  const [showImportNoiDungDropdown, setShowImportNoiDungDropdown] = useState(false);
  const [importNoiDungSearchTerm, setImportNoiDungSearchTerm] = useState('');
  const importNoiDungDropdownRef = useRef<HTMLDivElement>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const clickCountRef = useRef(0);
  const lastClickTimeRef = useRef(0);
  const projectNameRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    if (projectNameRef.current) {
      projectNameRef.current.style.height = 'auto';
      projectNameRef.current.style.height = `${projectNameRef.current.scrollHeight}px`;
    }
  }, [projectInfo["Tên dự án/công trình"], projectSearchTerm, showDropdown]);

  const handleLogoClick = () => {
    const now = Date.now();
    if (now - lastClickTimeRef.current < 500) {
      clickCountRef.current += 1;
    } else {
      clickCountRef.current = 1;
    }
    lastClickTimeRef.current = now;

    if (clickCountRef.current >= 5) {
      setShowSettings(prev => !prev);
      clickCountRef.current = 0;
    }
  };

  useEffect(() => {
    fetchUsers();
    
    // Handle click outside for custom dropdown
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setShowDropdown(false);
      }
      if (importProjectDropdownRef.current && !importProjectDropdownRef.current.contains(event.target as Node)) {
        setShowImportProjectDropdown(false);
      }
      if (importNoiDungDropdownRef.current && !importNoiDungDropdownRef.current.contains(event.target as Node)) {
        setShowImportNoiDungDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  useEffect(() => {
    if (currentUser) {
      fetchProjects(projectInfo);
      refreshCurrentProjectData(projectInfo["Mã công trình"] || "");
    }
  }, [currentUser]);

  const fetchUsers = async () => {
    try {
      const timestamp = new Date().getTime();
      const urlWithCacheBuster = `${USERS_GVIZ_URL}&t=${timestamp}`;
      let response;
      try {
        response = await fetch(urlWithCacheBuster);
        if (!response.ok) throw new Error('Direct fetch failed');
      } catch (e) {
        try {
          response = await fetch(`${USERS_PROXY_URL}&t=${timestamp}`);
          if (!response.ok) throw new Error('Proxy fetch failed');
        } catch (e2) {
          response = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(urlWithCacheBuster)}`);
        }
      }
      if (response && response.ok) {
        const csvText = await response.text();
        Papa.parse(csvText, {
          header: true,
          skipEmptyLines: true,
          complete: (results) => {
            setUsers(results.data);
          }
        });
      }
    } catch (err) {
      console.error("Failed to fetch users:", err);
    }
  };

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoggingIn(true);
    setLoginError(null);

    // Simulate a small delay for better UX
    setTimeout(() => {
      const user = users.find(u => u.username === loginForm.username && u.Password === loginForm.password);
      if (user) {
        const userData = {
          FullName: user.FullName,
          username: user.username,
          AllowedProjects: user.AllowedProjects || ''
        };
        setCurrentUser(userData);
        localStorage.setItem('currentUser', JSON.stringify(userData));
        setLoginForm({ username: '', password: '' });
      } else {
        setLoginError('Tên đăng nhập hoặc mật khẩu không chính xác.');
      }
      setIsLoggingIn(false);
    }, 800);
  };

  const [importForm, setImportForm] = useState({
    phongBan: 'QLĐT',
    tenCongTrinh: '',
    maCongTrinh: '',
    noiDung: '',
    soHieu: '',
    ngay: new Date().toISOString().split('T')[0],
  });
  const [importFile, setImportFile] = useState<File | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [importMessage, setImportMessage] = useState<{type: 'error' | 'success', text: string} | null>(null);
  const [showSuccessDialog, setShowSuccessDialog] = useState(false);
  const [duplicateWarning, setDuplicateWarning] = useState<{noiDung: string, soVb: string, ngayVb: string, fileUrl: string} | null>(null);
  const [rawDropdownData, setRawDropdownData] = useState<{phongBan: string, noiDung: string}[]>([]);

  // Fetch Dropdown sheet content
  useEffect(() => {
    const fetchDropdown = async () => {
      try {
        const timestamp = new Date().getTime();
        const DROPDOWN_GVIZ_URL = 'https://docs.google.com/spreadsheets/d/14BF0RUfBq-Arl6ngVvD44fQnNayBEC1Xtz-RFgzA4GI/gviz/tq?tqx=out:csv&sheet=' + encodeURIComponent('Dropdown');
        const urlWithCacheBuster = `${DROPDOWN_GVIZ_URL}&t=${timestamp}`;
        let response;
        try {
          response = await fetch(urlWithCacheBuster);
          if (!response.ok) throw new Error('Direct fetch failed');
        } catch (e) {
          try {
            response = await fetch(`https://corsproxy.io/?${encodeURIComponent(urlWithCacheBuster)}`);
            if (!response.ok) throw new Error('Proxy fetch failed');
          } catch (e2) {
            response = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(urlWithCacheBuster)}`);
          }
        }
        
        if (response && response.ok) {
          const text = await response.text();
          Papa.parse(text, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
              const rows = results.data as any[];
              const data: {phongBan: string, noiDung: string}[] = [];
              
              rows.forEach(row => {
                const vals = Object.values(row);
                // "Nội dung thả xuống" is usually Column B (index 1), "Tên sheet tạo dropdown" is Column C (index 2)
                const noiDung = row['Nội dung thả xuống'] || row['Noi dung tha xuong'] || vals[1]; 
                const phongBan = row['Tên sheet tạo dropdown'] || row['Ten sheet tao dropdown'] || vals[2];
                
                if (noiDung && typeof noiDung === 'string' && noiDung.trim() !== '') {
                  data.push({
                    phongBan: typeof phongBan === 'string' ? phongBan.trim() : '',
                    noiDung: noiDung.trim()
                  });
                }
              });
              
              setRawDropdownData(data);
            }
          });
        }
      } catch (e) {
        console.warn("Could not fetch Dropdown sheet", e);
      }
    };
    fetchDropdown();
  }, []);

  // Pre-fill the import form when modal opens
  useEffect(() => {
    if (showImportModal) {
      setImportForm(prev => ({
        ...prev,
        tenCongTrinh: projectInfo["Tên dự án/công trình"] || '',
        maCongTrinh: projectInfo["Mã công trình"] || ''
      }));
    }
  }, [showImportModal, projectInfo]);

  const handleImportSubmit = async (forceSave = false) => {
    if (!scriptUrl) {
      setImportMessage({ type: 'error', text: 'Vui lòng cấu hình Apps Script Web App URL.' });
      return;
    }
    
    if (!importForm.maCongTrinh || !importForm.phongBan || !importForm.noiDung) {
      setImportMessage({ type: 'error', text: 'Vui lòng điền mã công trình, phòng ban và nội dung.' });
      return;
    }

    setIsImporting(true);
    setImportMessage(null);
    setDuplicateWarning(null);

    if (!forceSave) {
      const duplicateInfo = await checkDuplicateDocument(importForm.maCongTrinh, importForm.phongBan, importForm.noiDung);
      if (duplicateInfo && (duplicateInfo.soVb || duplicateInfo.ngayVb || duplicateInfo.fileUrl)) {
        setIsImporting(false);
        setDuplicateWarning({
          noiDung: importForm.noiDung,
          soVb: duplicateInfo.soVb || '',
          ngayVb: duplicateInfo.ngayVb || '',
          fileUrl: duplicateInfo.fileUrl || ''
        });
        return;
      }
    }

    try {
      let base64File = '';
      let fileName = '';
      let mimeType = '';

      if (importFile) {
        const reader = new FileReader();
        reader.readAsDataURL(importFile);
        await new Promise((resolve, reject) => {
          reader.onload = () => resolve(null);
          reader.onerror = error => reject(error);
        });
        const result = reader.result as string;
        base64File = result.split(',')[1];
        fileName = importFile.name;
        mimeType = importFile.type;
      }

      // Pass the folder ID and data
      const payload = {
        action: 'importData',
        phongBan: importForm.phongBan,
        projectCode: importForm.maCongTrinh,
        noiDung: importForm.noiDung,
        soHieu: importForm.soHieu,
        ngay: importForm.ngay,
        fileName: fileName,
        mimeType: mimeType,
        fileData: base64File,
        targetFolderId: '13UFWLjjqJFc2omQ94TxwjJlMSl38Hwcj'
      };

      // Since Apps script blocks CORS responses on redirects, we use no-cors. 
      // We assume it succeeds if no network error is thrown.
      await fetch(scriptUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8',
        },
        body: JSON.stringify(payload),
        mode: 'no-cors'
      });

      // Show success, and schedule a data fetch
      setShowSuccessDialog(true);
      
      // Reset form variables to original states while preserving the active department
      setImportForm(prev => ({
        ...prev, 
        soHieu: '', 
        noiDung: '', 
        ngay: new Date().toISOString().split('T')[0],
        tenCongTrinh: '',
        maCongTrinh: ''
      }));
      setImportFile(null);
      setImportProjectSearchTerm('');
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
      
      // Polling for update (just refresh the current project data after a delay)
      if (projectInfo["Mã công trình"]) {
        setTimeout(() => refreshCurrentProjectData(projectInfo["Mã công trình"]), 2000);
      }
      
    } catch (err: any) {
      console.error(err);
      setImportMessage({ type: 'error', text: `Lỗi hệ thống: ${err.message}. Hãy cấp quyền DriveApp trong Apps Script.` });
    } finally {
      setIsImporting(false);
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem('currentUser');
    setProjectInfo({});
    setData([]);
  };

  const refreshCurrentProjectData = async (projectCode: string) => {
    if (!projectCode) {
      setLoading(false);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const { finalData, templateRows } = await loadProjectData(projectCode, templateCache || undefined);
      if (!templateCache) setTemplateCache(templateRows);
      setData(finalData);
    } catch(err: any) {
      setError(err.message || "Failed to refresh data");
    } finally {
      setLoading(false);
      setLastUpdated(new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
    }
  };

  const fetchProjects = async (currentProjectInfo: Record<string, string>) => {
    try {
      let response;
      const timestamp = new Date().getTime();
      const urlWithCacheBuster = `${PROJECTS_GVIZ_URL}&t=${timestamp}`;
      
      try {
        response = await fetch(urlWithCacheBuster);
        if (!response.ok) throw new Error('Direct fetch failed');
      } catch (e) {
        try {
          response = await fetch(`${PROJECTS_PROXY_URL}&t=${timestamp}`);
          if (!response.ok) throw new Error('Proxy fetch failed');
        } catch (e2) {
          try {
            response = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(urlWithCacheBuster)}`);
            if (!response.ok) throw new Error('AllOrigins fetch failed');
          } catch (e3) {
            // One last try
            response = await fetch(`https://thingproxy.freeboard.io/fetch/${encodeURIComponent(urlWithCacheBuster)}`);
            if (!response.ok) throw new Error('ThingProxy fetch failed');
          }
        }
      }
      
      const csvText = await response.text();
      Papa.parse(csvText, {
        header: false,
        skipEmptyLines: true,
        complete: (results) => {
          const rows = results.data as string[][];
          const projects: Record<string, string>[] = [];
          
          // The sheet "Thông tin theo MCT" structure:
          // Col 0: STT
          // Col 1: Mã công trình
          // Col 2: Tên công trình
          // Col 3: Địa điểm xây dựng
          // Col 4: Chủ đầu tư
          // Col 5: Đơn vị TV Thiết Kế
          // Col 6: Đơn vị TV Giám sát
          // Col 7: Đơn vị thi công
          
          // Start from row 1 (skipping header)
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const projectCode = row[1]?.trim() || "";
            
            // Filter projects based on user role
            if (currentUser) {
              const allowedProjects = currentUser.AllowedProjects.split(';').map(p => p.trim()).filter(p => p !== '');
              if (allowedProjects.length > 0 && !allowedProjects.includes(projectCode)) {
                continue;
              }
            }

            if (row && row.length >= 3 && row[2] && row[2].trim() !== "" && row[2] !== "Tên công trình") {
              projects.push({
                "Tên dự án/công trình": row[2].trim(),
                "Mã công trình": projectCode,
                "Chủ Đầu Tư": row[4]?.trim() || "",
                "Địa điểm xây dựng": row[3]?.trim() || "",
                "Đơn vị TV Thiết Kế": row[5]?.trim() || "",
                "Đơn vị TV Giám sát": row[6]?.trim() || "",
                "Đơn vị thi công": row[7]?.trim() || ""
              });
            }
          }
          
          // Ensure current project is in the list if not already
          const currentName = currentProjectInfo["Tên dự án/công trình"];
          if (currentName && !projects.some(p => p["Tên dự án/công trình"] === currentName)) {
            projects.unshift(currentProjectInfo);
          }
          
          // Remove duplicates by name
          const uniqueProjects = projects.filter((v, i, a) => 
            a.findIndex(t => t["Tên dự án/công trình"] === v["Tên dự án/công trình"]) === i
          );
          
          setAvailableProjects(uniqueProjects);
          
          if (uniqueProjects.length > 0 && (!currentProjectInfo || !currentProjectInfo["Mã công trình"])) {
            setProjectInfo(uniqueProjects[0]);
            refreshCurrentProjectData(uniqueProjects[0]["Mã công trình"]);
          }
        }
      });
    } catch (err) {
      console.error("Failed to fetch projects list:", err);
      setAvailableProjects([currentProjectInfo]);
    }
  };

  const saveScriptUrl = (url: string) => {
    setScriptUrl(url);
    localStorage.setItem('appsScriptUrl', url);
  };

  const triggerAction = async (actionName: string, actionId: string, params: Record<string, string> = {}) => {
    if (!scriptUrl) {
      alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
      setShowSettings(true);
      return;
    }

    setActionLoading(actionId);
    try {
      const url = new URL(scriptUrl);
      url.searchParams.append('action', actionId);
      
      // Append any additional parameters
      Object.entries(params).forEach(([key, value]) => {
        url.searchParams.append(key, value);
      });
      
      // We use no-cors because Apps Script Web Apps often have CORS issues unless configured properly.
      // With no-cors, the browser won't let us read the response, but the request DOES reach the server.
      await fetch(url.toString(), { mode: 'no-cors' });
      
      // Start polling almost immediately, as the Apps Script might be fast
      await new Promise(resolve => setTimeout(resolve, 100));
      
      // If we are updating project info, we expect the sheet to reflect this project
      if (actionId === 'updateProjectInfo' && params["Mã công trình"]) {
        // Use more frequent retries (30 attempts * 300ms = 9s max) to catch the update as soon as it happens
        refreshCurrentProjectData(params["Mã công trình"]);
      } else {
        refreshCurrentProjectData(projectInfo["Mã công trình"] || "");
      }
      
    } catch (err: any) {
      console.error(err);
      alert(`Lỗi khi thực hiện "${actionName}": Vui lòng kiểm tra lại kết nối hoặc URL Apps Script.`);
    } finally {
      setActionLoading(null);
    }
  };

  // Clean up data for display
  const displayData = data.filter(row => row.some(cell => cell.trim() !== ''));
  const headers = displayData.length > 0 ? displayData[0] : [];
  const allRows = displayData.length > 1 ? displayData.slice(1) : [];

  // Identify sections and assign rows to them
  const sections: string[] = [];
  const rowToSectionMap = new Map<number, string>();
  let currentSection = "";

  allRows.forEach((row, index) => {
    const isSectionHeader = row[1] && !row[3] && !row[4] && !row[5] && !row[6];
    if (isSectionHeader) {
      currentSection = row[1];
      if (!sections.includes(currentSection)) {
        sections.push(currentSection);
      }
    }
    rowToSectionMap.set(index, currentSection);
  });

  const rows = allRows.filter((row, index) => {
    const matchesSearch = row.some(cell => cell.toLowerCase().includes(searchTerm.toLowerCase()));
    const sectionOfRow = rowToSectionMap.get(index) || "";
    const matchesSection = selectedSection === 'all' || sectionOfRow === selectedSection;
    
    // If it's a section header, we always show it if it matches search OR if it's the selected section
    const isSectionHeader = row[1] && !row[3] && !row[4] && !row[5] && !row[6];
    if (isSectionHeader) {
      return (selectedSection === 'all' || row[1] === selectedSection) && matchesSearch;
    }

    return matchesSearch && matchesSection;
  });

  // Further filter rows to hide those in collapsed sections
  const visibleRows = rows.filter((row, index) => {
    const isSectionHeader = row[1] && !row[3] && !row[4] && !row[5] && !row[6];
    if (isSectionHeader) return true;
    
    const sectionOfRow = rowToSectionMap.get(allRows.indexOf(row));
    return !sectionOfRow || !collapsedSections.has(sectionOfRow);
  });

  const toggleSection = (sectionName: string) => {
    setCollapsedSections(prev => {
      const next = new Set(prev);
      if (next.has(sectionName)) {
        next.delete(sectionName);
      } else {
        next.add(sectionName);
      }
      return next;
    });
  };

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4 font-sans">
        <div className="max-w-md w-full bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-200">
          <div className="bg-white p-8 text-center border-b border-slate-100">
            <div className="inline-block mb-4 cursor-default select-none">
              <img 
                src={EVN_HCMC_LOGO} 
                alt="EVNHCMC" 
                className="h-24 w-auto object-contain"
                referrerPolicy="no-referrer"
              />
            </div>
            <h1 className="text-blue-900 text-2xl font-bold uppercase tracking-wider">Đăng nhập hệ thống</h1>
            <p className="text-slate-500 mt-2 text-sm font-medium">Danh mục ĐTXD</p>
          </div>
          
          <form onSubmit={handleLogin} className="p-8 space-y-6">
            {loginError && (
              <div className="bg-red-50 border-l-4 border-red-500 p-4 flex gap-3 items-start animate-in fade-in slide-in-from-top-2">
                <AlertCircle className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
                <p className="text-sm text-red-700">{loginError}</p>
              </div>
            )}
            
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-1.5">Tên đăng nhập</label>
                <div className="relative">
                  <UserIcon className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                  <input 
                    type="text"
                    required
                    value={loginForm.username}
                    onChange={(e) => setLoginForm(prev => ({ ...prev, username: e.target.value }))}
                    className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all text-slate-800"
                    placeholder="Nhập tên đăng nhập..."
                  />
                </div>
              </div>
              
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-1.5">Mật khẩu</label>
                <div className="relative">
                  <Lock className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                  <input 
                    type="password"
                    required
                    value={loginForm.password}
                    onChange={(e) => setLoginForm(prev => ({ ...prev, password: e.target.value }))}
                    className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all text-slate-800"
                    placeholder="Nhập mật khẩu..."
                  />
                </div>
              </div>
            </div>
            
            <button 
              type="submit"
              disabled={isLoggingIn}
              className="w-full bg-blue-900 text-white py-3.5 rounded-xl font-bold uppercase tracking-widest hover:bg-blue-800 active:scale-[0.98] transition-all shadow-lg shadow-blue-900/20 flex items-center justify-center gap-2 disabled:opacity-70"
            >
              {isLoggingIn ? (
                <RefreshCw className="w-5 h-5 animate-spin" />
              ) : (
                <LogIn className="w-5 h-5" />
              )}
              <span>Đăng nhập</span>
            </button>
          </form>
          
          <div className="px-8 py-4 bg-slate-50 border-t border-slate-100 text-center">
            <p className="text-xs text-slate-400">© 2026 EVNHCMC - Hệ thống quản lý nội bộ</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex items-center cursor-pointer select-none" onClick={handleLogoClick} title="EVNHCMC">
              <img 
                src={EVN_HCMC_LOGO} 
                alt="EVNHCMC" 
                className="h-16 w-auto object-contain"
                referrerPolicy="no-referrer"
              />
            </div>
            <h1 className="text-xl sm:text-2xl font-bold text-blue-900 uppercase tracking-tight">DANH MỤC ĐTXD</h1>
          </div>
          <div className="flex items-center gap-4">
            <div className="hidden sm:flex flex-col items-end mr-2">
              <span className="text-sm font-bold text-slate-700">{currentUser.FullName}</span>
              <span className="text-[10px] text-slate-400 uppercase tracking-wider font-medium">@{currentUser.username}</span>
            </div>
            <button 
              onClick={handleLogout}
              className="p-2.5 bg-slate-100 text-slate-600 rounded-full hover:bg-red-50 hover:text-red-600 transition-all border border-slate-200"
              title="Đăng xuất"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Import Modal */}
        {showImportModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
            <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg flex flex-col max-h-[90vh]">
              <div className="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-white rounded-t-2xl">
                <div className="flex items-center gap-2">
                  <Database className="w-5 h-5 text-blue-600" />
                  <h3 className="font-bold text-slate-800 uppercase text-lg">Thêm mới văn bản</h3>
                </div>
                <button 
                  onClick={() => {
                    setShowImportModal(false);
                    setImportMessage(null);
                    setShowSuccessDialog(false);
                    if (projectInfo["Mã công trình"]) {
                      refreshCurrentProjectData(projectInfo["Mã công trình"]);
                    }
                  }}
                  className="p-2 hover:bg-slate-100 rounded-full text-slate-400 hover:text-slate-600 transition-all"
                  title="Đóng cửa sổ"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>
              
              <div className="p-6 overflow-y-auto flex-1 bg-white">
                <div className="space-y-5">
                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Chọn Phòng ban/Đơn vị:</label>
                    <select
                      value={importForm.phongBan}
                      onChange={e => setImportForm({...importForm, phongBan: e.target.value})}
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm bg-white"
                    >
                      {["KTAT","KHVT","QLĐT","TCG","TVTK","TVGS","XL","TCKT","QLLĐ","DVĐL","ĐĐHTĐ","ETC","Kiểm toán","SXD","TCT","UBND"].map(dept => (
                        <option key={dept} value={dept}>{dept}</option>
                      ))}
                    </select>
                  </div>

                  <div className="relative z-[50]" ref={importProjectDropdownRef}>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Tên công trình:</label>
                    <input
                      type="text"
                      value={showImportProjectDropdown ? importProjectSearchTerm : importForm.tenCongTrinh}
                      onChange={e => setImportProjectSearchTerm(e.target.value)}
                      onFocus={() => {
                        setImportProjectSearchTerm('');
                        setShowImportProjectDropdown(true);
                      }}
                      placeholder="Nhập tên hoặc mã CT để tìm kiếm..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm"
                    />
                    {showImportProjectDropdown && (
                      <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-xl max-h-60 overflow-y-auto z-[60]">
                        {availableProjects.filter(p => !importProjectSearchTerm || (p["Tên dự án/công trình"]?.toLowerCase().includes(importProjectSearchTerm.toLowerCase()) || p["Mã công trình"]?.toLowerCase().includes(importProjectSearchTerm.toLowerCase()))).map((project, idx) => (
                          <div
                            key={idx}
                            onClick={() => {
                              setImportForm({
                                ...importForm, 
                                tenCongTrinh: project["Tên dự án/công trình"],
                                maCongTrinh: project["Mã công trình"]
                              });
                              setShowImportProjectDropdown(false);
                            }}
                            className="px-4 py-3 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0 transition-colors"
                          >
                            <div className="font-medium text-slate-800">{project["Tên dự án/công trình"]}</div>
                            <div className="text-xs text-slate-500 mt-1">Mã CT: {project["Mã công trình"]}</div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Mã công trình:</label>
                    <input
                      type="text"
                      value={importForm.maCongTrinh}
                      disabled
                      placeholder="Mã công trình sẽ tự động điền..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-500 shadow-sm bg-slate-100 cursor-not-allowed"
                    />
                  </div>

                  <div className="relative z-[40]" ref={importNoiDungDropdownRef}>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Nội dung văn bản:</label>
                    <input
                      type="text"
                      value={showImportNoiDungDropdown ? importNoiDungSearchTerm : importForm.noiDung}
                      onChange={e => setImportNoiDungSearchTerm(e.target.value)}
                      onFocus={() => {
                        setImportNoiDungSearchTerm('');
                        setShowImportNoiDungDropdown(true);
                      }}
                      placeholder="Nhập nội dung để tìm kiếm..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm bg-white"
                    />
                    {showImportNoiDungDropdown && (
                      <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[60]">
                        {Array.from(new Set(rawDropdownData.filter(item => item.phongBan === importForm.phongBan).map(item => item.noiDung)))
                          .filter(item => !importNoiDungSearchTerm || String(item).toLowerCase().includes(importNoiDungSearchTerm.toLowerCase()))
                          .map((item, idx) => (
                            <div
                              key={idx}
                              onClick={() => {
                                setImportForm({...importForm, noiDung: item});
                                setShowImportNoiDungDropdown(false);
                              }}
                              className="px-4 py-2 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0 transition-colors text-sm text-slate-700"
                            >
                              {item}
                            </div>
                          ))}
                        {Array.from(new Set(rawDropdownData.filter(item => item.phongBan === importForm.phongBan).map(item => item.noiDung)))
                          .filter(item => !importNoiDungSearchTerm || String(item).toLowerCase().includes(importNoiDungSearchTerm.toLowerCase())).length === 0 && (
                          <div className="px-4 py-3 text-sm text-slate-400 italic">Không tìm thấy nội dung phù hợp</div>
                        )}
                      </div>
                    )}
                  </div>
                  
                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Số hiệu văn bản:</label>
                    <input
                      type="text"
                      value={importForm.soHieu}
                      onChange={e => setImportForm({...importForm, soHieu: e.target.value})}
                      placeholder="Nhập số văn bản..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm"
                    />
                  </div>

                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Ngày văn bản:</label>
                    <input
                      type="date"
                      value={importForm.ngay}
                      onChange={e => setImportForm({...importForm, ngay: e.target.value})}
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm"
                    />
                  </div>

                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Đính kèm văn bản (PDF, Ảnh, Word...):</label>
                    <input
                      type="file"
                      ref={fileInputRef}
                      onChange={e => {
                        if (e.target.files && e.target.files.length > 0) {
                          setImportFile(e.target.files[0]);
                        } else {
                          setImportFile(null);
                        }
                      }}
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 bg-white shadow-sm file:mr-4 file:py-1 file:px-3 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-slate-100 file:text-slate-700 hover:file:bg-slate-200 cursor-pointer"
                    />
                  </div>

                  <button
                    onClick={() => handleImportSubmit(false)}
                    disabled={isImporting}
                    className="w-full py-2.5 bg-blue-600 text-white rounded-md hover:bg-blue-700 font-bold transition-colors disabled:bg-blue-400 mt-2 shadow-sm"
                  >
                    {isImporting ? 'ĐANG LƯU...' : 'LƯU'}
                  </button>

                  {importMessage && importMessage.type === 'error' && (
                    <div className="mt-4 text-sm font-bold flex items-center justify-center gap-2 text-red-600">
                      <X className="w-5 h-5 font-bold" />
                      {importMessage.text}
                    </div>
                  )}
                </div>
              </div>
              <div className="px-6 py-4 border-t border-slate-200 flex justify-end bg-white rounded-b-2xl">
                <button 
                  onClick={() => {
                    setShowImportModal(false);
                    setImportMessage(null);
                    setShowSuccessDialog(false);
                    if (projectInfo["Mã công trình"]) {
                      refreshCurrentProjectData(projectInfo["Mã công trình"]);
                    }
                  }}
                  className="px-6 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 font-bold transition-colors shadow-sm"
                >
                  HOÀN TẤT
                </button>
              </div>

              {/* Duplicate Warning Dialog Overlay */}
              {duplicateWarning && (
                <div className="absolute inset-0 z-[70] flex items-center justify-center bg-white/70 backdrop-blur-sm rounded-2xl p-4">
                  <div className="bg-white p-6 rounded-xl shadow-2xl border border-amber-200 flex flex-col w-full max-w-md mx-auto animate-in fade-in zoom-in duration-200">
                    <div className="flex items-center gap-3 mb-4">
                      <div className="w-10 h-10 bg-amber-100 text-amber-600 rounded-full flex items-center justify-center shrink-0">
                        <AlertCircle className="w-6 h-6" />
                      </div>
                      <h3 className="text-lg font-bold text-slate-800">Cảnh báo: Đã có văn bản!</h3>
                    </div>
                    
                    <p className="text-sm text-slate-600 mb-4">
                      Một văn bản có cùng nội dung đã được nhập trước đó. Thông tin chi tiết:
                    </p>
                    
                    <div className="bg-slate-50 p-4 rounded-lg border border-slate-200 space-y-3 text-sm mb-6">
                      <div>
                        <span className="text-slate-500 font-semibold block mb-0.5">Nội dung văn bản:</span>
                        <span className="text-slate-800 font-medium">{duplicateWarning.noiDung}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-4">
                        <div>
                          <span className="text-slate-500 font-semibold block mb-0.5">Số VB:</span>
                          <span className="text-slate-800 font-medium">{duplicateWarning.soVb || <span className="italic text-slate-400">Không có</span>}</span>
                        </div>
                        <div>
                          <span className="text-slate-500 font-semibold block mb-0.5">Ngày VB:</span>
                          <span className="text-slate-800 font-medium">{duplicateWarning.ngayVb || <span className="italic text-slate-400">Không có</span>}</span>
                        </div>
                      </div>
                      <div>
                        <span className="text-slate-500 font-semibold block mb-0.5">File đính kèm:</span>
                        {duplicateWarning.fileUrl && duplicateWarning.fileUrl.startsWith('http') ? (
                          <a href={duplicateWarning.fileUrl} target="_blank" rel="noopener noreferrer" className="text-indigo-600 hover:underline inline-flex items-center gap-1">
                            <Eye className="w-3.5 h-3.5" /> Xem file
                          </a>
                        ) : (
                          <span className="italic text-slate-400">{duplicateWarning.fileUrl || 'Không có'}</span>
                        )}
                      </div>
                    </div>
                    
                    <div className="flex gap-3">
                      <button 
                        onClick={() => setDuplicateWarning(null)}
                        className="flex-1 py-2.5 bg-slate-100 text-slate-700 rounded-lg hover:bg-slate-200 font-bold transition-colors"
                      >
                        Hủy bỏ
                      </button>
                      <button 
                        onClick={() => handleImportSubmit(true)}
                        className="flex-1 py-2.5 bg-amber-500 text-white rounded-lg hover:bg-amber-600 font-bold transition-colors shadow-sm"
                      >
                        Vẫn tiếp tục lưu
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* Success Dialog Overlay */}
              {showSuccessDialog && (
                <div className="absolute inset-0 z-50 flex items-center justify-center bg-white/60 backdrop-blur-sm rounded-2xl">
                  <div className="bg-white p-6 rounded-xl shadow-xl border border-emerald-100 flex flex-col items-center justify-center text-center max-w-sm w-full mx-4 animate-in fade-in zoom-in duration-200">
                    <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mb-4">
                      <svg className="w-8 h-8" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={3}>
                        <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                      </svg>
                    </div>
                    <h3 className="text-xl font-bold text-slate-800 mb-2">Lưu thành công!</h3>
                    <p className="text-sm text-slate-500 mb-6">Dữ liệu văn bản đã được cập nhật thành công.</p>
                    <button 
                      onClick={() => setShowSuccessDialog(false)}
                      className="w-full py-2.5 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 font-bold transition-colors"
                    >
                      Đóng
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Settings Panel */}
        {showSettings && (
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 mb-8 animate-in slide-in-from-top-4 fade-in duration-200">
            <h2 className="text-lg font-medium mb-4 flex items-center gap-2">
              <Settings className="w-5 h-5 text-slate-500" />
              Cấu hình kết nối Google Apps Script
            </h2>
            <div className="bg-amber-50 border border-amber-200 rounded-lg p-4 mb-4 text-sm text-amber-800 flex gap-3">
              <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
              <div>
                <p className="font-medium mb-1">Tại sao cần cấu hình URL này?</p>
                <p className="mb-2">Để các nút chức năng hoạt động, bạn cần triển khai mã Apps Script trên Google Sheet của bạn dưới dạng <strong>Web App</strong> và dán URL vào đây.</p>
                <ol className="list-decimal ml-5 space-y-1">
                  <li>Mở Google Sheet, vào <strong>Tiện ích mở rộng &gt; Apps Script</strong>.</li>
                  <li>Thêm hàm <code>doGet(e)</code> để xử lý các tham số <code>action</code> ('create', 'import', 'export').</li>
                  <li>Chọn <strong>Triển khai &gt; Triển khai mới</strong>, chọn loại <strong>Ứng dụng web</strong>.</li>
                  <li>Quyền truy cập: <strong>Bất kỳ ai</strong>.</li>
                  <li>Sao chép URL Web App và dán vào ô bên dưới.</li>
                </ol>
              </div>
            </div>
            <div className="flex gap-3">
              <input 
                type="url" 
                value={scriptUrl}
                onChange={(e) => saveScriptUrl(e.target.value)}
                placeholder="https://script.google.com/macros/s/.../exec"
                className="flex-1 px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none transition-shadow"
              />
              <button 
                onClick={() => setShowSettings(false)}
                className="px-4 py-2 bg-slate-800 text-white rounded-lg hover:bg-slate-700 font-medium transition-colors"
              >
                Lưu & Đóng
              </button>
            </div>
          </div>
        )}

        {/* Action Buttons */}
        <div className="hidden sm:block mb-8 bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-slate-200">
          <div className="grid grid-cols-2 sm:flex items-center gap-3 sm:gap-4 w-full sm:w-auto">
            <button 
              onClick={() => {
                if (!scriptUrl) {
                  alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
                  setShowSettings(true);
                  return;
                }
                setShowImportModal(true);
              }}
              className="group relative overflow-hidden flex flex-col sm:flex-row items-center justify-center gap-1.5 sm:gap-2 px-3 py-2.5 sm:px-8 sm:py-3.5 bg-gradient-to-br from-blue-600 to-blue-700 text-white rounded-xl font-bold shadow-lg shadow-blue-500/20 ring-1 ring-inset ring-white/20 hover:shadow-xl hover:shadow-blue-500/40 hover:-translate-y-0.5 active:scale-95 transition-all duration-300"
            >
              <div className="absolute inset-0 w-1/2 bg-gradient-to-r from-transparent via-white/40 to-transparent group-hover:animate-shine opacity-0 group-hover:opacity-100 transition-opacity" />
              <FilePlus className="w-4 h-4 sm:w-5 sm:h-5 relative z-10" />
              <span className="uppercase tracking-wider text-[10px] sm:text-sm relative z-10">Thêm mới</span>
            </button>
            
            <button 
              onClick={() => {
                try {
                  const wb = XLSX.utils.book_new();
                  const wsData: any[][] = [];
                  const merges: any[] = [];
                  
                  // Thêm 2 hàng đầu tiên cho Tiêu đề trang tính
                  wsData.push(['DANH MỤC HỒ SƠ', '', '', '', '', '', '', '']); // Dòng 1
                  wsData.push(['', '', '', '', '', '', '', '']);                 // Dòng 2
                  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } });         // Merge A1:H1
                  
                  let rowIndex = 2;
                  
                  // Phần mảng thông tin dự án (projectInfo) được bóc tách và chèn vào các hàng tiếp theo
                  Object.entries(projectInfo).forEach(([key, value]) => {
                    wsData.push([`${key}:`, '', value, '', '', '', '', '']); // 8 cột
                    merges.push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: 1 } });
                    merges.push({ s: { r: rowIndex, c: 2 }, e: { r: rowIndex, c: 3 } }); // merge C:D
                    rowIndex++;
                  });
                  
                  // Bù thêm các dòng trống nếu thông tin dự án ít hơn để Header chính xác ở dòng 11 (rowIndex 10)
                  while (rowIndex < 9) {
                     wsData.push(['', '', '', '', '', '', '', '']);
                     merges.push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: 1 } });
                     merges.push({ s: { r: rowIndex, c: 2 }, e: { r: rowIndex, c: 3 } }); // merge C:D
                     rowIndex++;
                  }
                  
                  // Dòng 10 (trống) trước khi vào header
                  wsData.push(['', '', '', '', '', '', '', '']);
                  rowIndex++;
                  
                  const headerRowIndex = rowIndex; // = 10 (hiển thị là số 11 trên excel)
                  wsData.push(['STT', 'TÊN VĂN BẢN', '', 'SỐ VB', 'NGÀY VB', 'FILE VB', 'CƠ QUAN BAN HÀNH', 'GHI CHÚ']);
                  merges.push({ s: { r: headerRowIndex, c: 1 }, e: { r: headerRowIndex, c: 2 } });
                  
                  const borderStyle = { style: 'thin', color: { rgb: "000000" } };
                  const borderObj = { top: borderStyle, bottom: borderStyle, left: borderStyle, right: borderStyle };
                  
                  // Xử lý dữ liệu bảng (data) (bỏ qua hàng tiêu đề)
                  const exportData = data.length > 0 ? data.slice(1) : [];
                  exportData.forEach((row) => {
                     // row[5] là cột gán đường link, vd: 'https://...'
                     let fileLink = row[5];
                     let hasLink = fileLink && (fileLink.startsWith('http://') || fileLink.startsWith('https://'));
                     let fileDisp = hasLink ? "Xem File" : (fileLink || "");
                     
                     wsData.push([
                       { v: row[0] || '', t: 's' }, 
                       { v: row[1] || '', t: 's' }, 
                       { v: '', t: 's' }, 
                       { v: row[3] || '', t: 's' }, 
                       { v: row[4] || '', t: 's' }, 
                       { v: fileDisp, t: 's' }, 
                       { v: row[6] || '', t: 's' }, 
                       { v: '', t: 's' } // For "Ghi chú"
                     ]);
                  });
                  
                  const ws = XLSX.utils.aoa_to_sheet(wsData);
                  
                  // Gộp ô TÊN VĂN BẢN (từ dòng Header trở xuống)
                  for(let r = headerRowIndex + 1; r < wsData.length; r++) {
                     merges.push({ s: { r: r, c: 1 }, e: { r: r, c: 2 } });
                  }
                  
                  ws['!merges'] = merges;
                  
                  // Định đạng độ rộng các cột
                  ws['!cols'] = [
                    { wch: 8 },   // A: STT
                    { wch: 20 },  // B: TÊN VĂN BẢN (merge C)
                    { wch: 40 },  // C: part of TÊN VĂN BẢN / project value 
                    { wch: 20 },  // D: SỐ VB
                    { wch: 15 },  // E: NGÀY VB
                    { wch: 15 },  // F: FILE VB
                    { wch: 15 },  // G: CQBH
                    { wch: 20 },  // H: GHI CHÚ
                  ];
                  
                  // Cài đặt font cho Info Key & Info Value (các dòng trước Header)
                  for (let r = 0; r < headerRowIndex; r++) {
                    for (let c = 0; c <= 7; c++) {
                       const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
                       if (!ws[cellRef]) ws[cellRef] = { v: '', t: 's' };
                       ws[cellRef].s = {
                         font: { name: 'Times New Roman', sz: 11 },
                         alignment: { vertical: 'center' }
                       };
                       
                       if (r === 0 && c === 0) { // Định dạng tiêu đề chính "DANH MỤC HỒ SƠ"
                           ws[cellRef].s.font.bold = true;
                           ws[cellRef].s.font.sz = 14;
                           ws[cellRef].s.alignment.horizontal = 'center';
                       } else if ((c === 0 || c === 2) && r >= 2) {
                         ws[cellRef].s.alignment.wrapText = true;
                       }
                    }
                  }

                  // Áp dụng format (Border, Center, Bold cho Header, Wrap cho B & C)
                  for (let r = headerRowIndex; r < wsData.length; r++) {
                    for (let c = 0; c <= 7; c++) {
                       const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
                       if (!ws[cellRef]) ws[cellRef] = { v: '', t: 's' };
                       ws[cellRef].s = {
                         border: borderObj,
                         font: { name: 'Times New Roman', sz: 11 },
                         alignment: { vertical: 'center' }
                       };
                       
                       // Ensure text format to prevent Excel from mangling fractions or dates
                       ws[cellRef].z = '@';
                       
                       if (r === headerRowIndex) { // Đoạn này định dạng cho dòng Header
                         ws[cellRef].s.font.bold = true;
                         ws[cellRef].s.alignment.horizontal = 'center';
                       } else { // Phía dữ liệu
                         if (c === 0 || c === 3 || c === 4 || c === 5 || c === 6) { // Thêm cột 6 (G)
                            ws[cellRef].s.alignment.horizontal = 'center';
                            if (c === 3) {
                              ws[cellRef].s.alignment.wrapText = true; // Cho phép xuống dòng ở Số VB
                            }
                         } else if (c === 1 || c === 2) {
                            ws[cellRef].s.alignment.wrapText = true;
                         }
                       }
                    }
                  }
                  
                  // Cấu hình gắn hyperlink cho dòng file (File VB)
                  exportData.forEach((row, idx) => {
                     let fileLink = row[5];
                     if (fileLink && (fileLink.startsWith('http://') || fileLink.startsWith('https://'))) {
                        const r = headerRowIndex + 1 + idx;
                        const cellRef = XLSX.utils.encode_cell({ r: r, c: 5 }); // Cột F
                        if (ws[cellRef]) {
                           ws[cellRef].l = { Target: fileLink };
                           ws[cellRef].s.font.color = { rgb: '0000FF' };
                           ws[cellRef].s.font.underline = true;
                        }
                     }
                  });

                  // Tạo Sheet vào workbook và lưu
                  XLSX.utils.book_append_sheet(wb, ws, "Danh mục hồ sơ");
                  const projectCode = projectInfo["Mã công trình"] || "Unknown";
                  const filename = `Danh mục hồ sơ - ${projectCode}.xlsx`;
                  XLSX.writeFile(wb, filename);
                  
                } catch (error) {
                  console.error('Error generating Excel file:', error);
                  alert('Có lỗi xảy ra khi tạo file Excel. Vui lòng thử lại.');
                }
              }}
              className="group relative overflow-hidden flex flex-col sm:flex-row items-center justify-center gap-1.5 sm:gap-2 px-3 py-2.5 sm:px-8 sm:py-3.5 bg-white text-emerald-600 border border-emerald-200 rounded-xl font-bold shadow-md shadow-slate-200/50 hover:shadow-xl hover:shadow-emerald-500/20 hover:border-emerald-300 hover:-translate-y-0.5 active:scale-95 transition-all duration-300 cursor-pointer"
            >
              <div className="absolute inset-0 w-1/2 bg-gradient-to-r from-transparent via-emerald-100/50 to-transparent group-hover:animate-shine opacity-0 group-hover:opacity-100 transition-opacity" />
              <Download className="w-4 h-4 sm:w-5 sm:h-5 relative z-10" />
              <span className="uppercase tracking-wider text-[10px] sm:text-sm relative z-10">Xuất Excel</span>
            </button>
          </div>
        </div>

        {/* Project Info */}
        {Object.keys(projectInfo).length > 0 && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 mb-8">
            <h2 className="text-lg font-semibold text-slate-800 mb-4 border-b border-slate-100 pb-2">Thông tin dự án</h2>
            <div className="flex flex-col gap-y-3">
              {Object.entries(projectInfo).map(([key, value], idx) => {
                if (key === 'Tên dự án/công trình') {
                  return (
                    <div key={idx} className="flex flex-col sm:flex-row sm:items-start gap-1 sm:gap-3">
                      <span className="text-slate-500 font-medium min-w-[160px] pt-1.5">{key}:</span>
                      <div className="relative z-30 flex-1 max-w-4xl" ref={dropdownRef}>
                        <textarea 
                          ref={projectNameRef}
                          value={showDropdown ? projectSearchTerm : value}
                          onFocus={() => {
                            setProjectSearchTerm('');
                            setShowDropdown(true);
                          }}
                          onInput={(e) => {
                            const target = e.target as HTMLTextAreaElement;
                            target.style.height = 'auto';
                            target.style.height = `${target.scrollHeight}px`;
                          }}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                            }
                          }}
                          onChange={(e) => {
                            setProjectSearchTerm(e.target.value);
                          }}
                          className="w-full pl-3 pr-10 py-1.5 border border-slate-300 rounded-md focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none text-slate-800 font-semibold resize-none overflow-hidden"
                          placeholder="Nhập từ khoá để tìm kiếm..."
                          rows={1}
                        />
                        {((showDropdown && projectSearchTerm) || (!showDropdown && value)) && (
                          <button 
                            onClick={() => {
                              if (showDropdown) {
                                setProjectSearchTerm('');
                              } else {
                                setProjectInfo(prev => ({...prev, "Tên dự án/công trình": ""}));
                                setProjectSearchTerm('');
                                setShowDropdown(true);
                              }
                            }}
                            className="absolute right-3 top-2 p-1 text-slate-400 hover:text-slate-600"
                            title="Xoá"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        )}
                        
                        {showDropdown && (
                          <div className="absolute z-50 left-0 right-0 mt-1 bg-white border border-slate-200 rounded-md shadow-lg max-h-80 overflow-y-auto">
                            {availableProjects
                              .filter(p => !projectSearchTerm || (String(p["Tên dự án/công trình"]).toLowerCase().includes(projectSearchTerm.toLowerCase()) || String(p["Mã công trình"] || '').toLowerCase().includes(projectSearchTerm.toLowerCase())))
                              .map((p, i) => (
                                <div 
                                  key={i}
                                  className="px-4 py-2 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0 leading-relaxed"
                                  onClick={async () => {
                                    setProjectInfo(p);
                                    setData([]);
                                    setLoading(true);
                                    setShowDropdown(false);
                                    setProjectSearchTerm('');
                                    try {
                                      const { finalData, templateRows } = await loadProjectData(p["Mã công trình"], templateCache || undefined);
                                      if (!templateCache) setTemplateCache(templateRows);
                                      setData(finalData);
                                    } catch(e: any) {
                                      setError(e.message || "Failed to load project data");
                                    } finally {
                                      setLoading(false);
                                      setLastUpdated(new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
                                    }
                                  }}
                                >
                                  <div className="font-medium text-sm text-slate-700">{p["Tên dự án/công trình"]}</div>
                                  {p["Mã công trình"] && (
                                    <div className="text-xs text-slate-500 mt-1">Mã CT: {p["Mã công trình"]}</div>
                                  )}
                                </div>
                              ))}
                            {availableProjects.filter(p => !projectSearchTerm || (String(p["Tên dự án/công trình"]).toLowerCase().includes(projectSearchTerm.toLowerCase()) || String(p["Mã công trình"] || '').toLowerCase().includes(projectSearchTerm.toLowerCase()))).length === 0 && (
                              <div className="px-4 py-3 text-sm text-slate-400 italic">Không tìm thấy dự án phù hợp</div>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  );
                }
                return (
                  <div key={idx} className="flex flex-col sm:flex-row sm:items-start gap-1 sm:gap-3">
                    <span className="text-slate-500 font-medium min-w-[160px]">{key}:</span>
                    <span className="text-slate-800 font-semibold">{value}</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Data Table */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
          <div className="px-6 py-4 border-b border-slate-200 bg-white flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div className="flex justify-between items-center gap-4 w-full lg:w-auto">
              <h2 className="font-semibold text-slate-800 shrink-0 uppercase tracking-wider text-sm">DANH MỤC HỒ SƠ</h2>
              
              <div className="flex items-center gap-2 lg:hidden">
                <span className="text-[10px] font-medium text-slate-500 bg-slate-100 px-2 py-1 rounded-full whitespace-nowrap">
                  {visibleRows.length} dòng
                </span>
                <button 
                  onClick={() => setShowMobileFilters(!showMobileFilters)}
                  className={`p-2 rounded-lg border transition-all relative ${
                    showMobileFilters || selectedSection !== 'all' 
                      ? 'bg-emerald-50 border-emerald-200 text-emerald-600' 
                      : 'bg-white border-slate-200 text-slate-400 hover:text-slate-600'
                  }`}
                  title="Lọc theo mục"
                >
                  <Filter className="w-4 h-4" />
                  {selectedSection !== 'all' && (
                    <span className="absolute top-1 right-1 w-2 h-2 bg-emerald-500 rounded-full border border-white"></span>
                  )}
                </button>
              </div>
            </div>

            {/* Search and Filters Area */}
            <div className={`${showMobileFilters ? 'flex' : 'hidden'} lg:flex flex-col lg:flex-row lg:items-center gap-3 w-full lg:w-auto`}>
              {/* Section Filter - Toggleable on mobile, always visible on desktop */}
              <div className="relative w-full lg:w-64">
                <select
                  value={selectedSection}
                  onChange={(e) => setSelectedSection(e.target.value)}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none bg-white appearance-none pr-10"
                >
                  <option value="all">Tất cả các mục</option>
                  {sections.map((section, idx) => (
                    <option key={idx} value={section}>{section}</option>
                  ))}
                </select>
                <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" />
              </div>

              {/* Search - Always visible on desktop, toggleable with filters on mobile if needed, but user wanted it outside on mobile before */}
              {/* Actually, let's keep search always visible as per previous request, but align it nicely */}
              <div className="relative w-full lg:w-64">
                <input
                  type="text"
                  placeholder="Tìm kiếm nội dung..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="px-3 pr-9 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none w-full"
                />
                {searchTerm && (
                  <button 
                    onClick={() => setSearchTerm('')}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 p-1 text-slate-400 hover:text-slate-600 transition-colors"
                    title="Xoá tìm kiếm"
                  >
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>

              <span className="hidden lg:inline-flex text-xs font-medium text-slate-500 bg-slate-200 px-2.5 py-1 rounded-full whitespace-nowrap">
                {visibleRows.length} dòng
              </span>
            </div>
          </div>
          
          {loading ? (
            <div className="p-12 flex flex-col items-center justify-center text-slate-500">
              <RefreshCw className="w-8 h-8 animate-spin mb-4 text-emerald-600" />
              <p>Đang tải dữ liệu ...</p>
            </div>
          ) : error ? (
            <div className="p-8 text-center text-red-600">
              <AlertCircle className="w-12 h-12 mx-auto mb-3 opacity-50" />
              <p className="font-medium">Lỗi khi tải dữ liệu</p>
              <p className="text-sm mt-1">{error}</p>
            </div>
          ) : (
            <>
              {/* Desktop View */}
              <div className="hidden md:block overflow-x-auto max-h-[600px]">
                <table className="w-full text-sm text-left whitespace-nowrap">
                  <thead className="text-xs text-slate-600 uppercase bg-white sticky top-0 z-10 shadow-sm text-center">
                    <tr>
                      {/* Custom headers based on the CSV structure we saw */}
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">STT</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">TÊN VĂN BẢN</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200 w-32 max-w-[160px]">SỐ VB</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">NGÀY VB</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">FILE VB</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">CƠ QUAN BAN HÀNH</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {visibleRows.map((row, rowIndex) => {
                      // Skip empty rows or rows that are just section headers if they don't fit well
                      // But we'll try to render them nicely
                      const isSectionHeader = row[1] && !row[3] && !row[4] && !row[5] && !row[6];
                      const isCollapsed = isSectionHeader && collapsedSections.has(row[1]);
                      
                      return (
                        <tr 
                          key={rowIndex} 
                          className={`hover:bg-slate-50 transition-colors ${isSectionHeader ? 'bg-slate-100 font-semibold text-emerald-800 cursor-pointer' : ''}`}
                          onClick={() => isSectionHeader && toggleSection(row[1])}
                        >
                          <td className="px-4 py-3 border-r border-slate-100 text-center text-slate-500">
                            {isSectionHeader ? (
                              <div className="flex items-center justify-center">
                                <ChevronDown className={`w-4 h-4 transition-transform duration-200 ${isCollapsed ? '-rotate-90' : ''}`} />
                              </div>
                            ) : row[0]}
                          </td>
                          <td className="px-4 py-3 border-r border-slate-100 whitespace-normal min-w-[300px] text-slate-700">{row[1]}</td>
                          <td className="px-4 py-3 border-r border-slate-100 text-center whitespace-normal w-32 max-w-[160px] break-words">{row[3]}</td>
                          <td className="px-4 py-3 border-r border-slate-100">{row[4]}</td>
                          <td className="px-4 py-3 border-r border-slate-100">
                            {row[5] && (row[5].startsWith('http://') || row[5].startsWith('https://')) ? (
                              <a 
                                href={row[5]} 
                                target="_blank" 
                                rel="noopener noreferrer"
                                className="inline-flex items-center gap-1.5 text-indigo-600 hover:text-indigo-800 hover:bg-indigo-50 px-2 py-1 rounded-md transition-all font-medium group"
                              >
                                <Eye className="w-4 h-4 group-hover:scale-110 transition-transform" />
                                <span>Xem file</span>
                              </a>
                            ) : (
                              <span className={row[5]?.trim().toLowerCase() === 'xem file' ? 'text-slate-400 italic' : ''}>
                                {row[5]}
                              </span>
                            )}
                          </td>
                          <td className="px-4 py-3 border-r border-slate-100 text-center">{row[6]}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Mobile View (Module/Card Layout) */}
              <div className="md:hidden divide-y divide-slate-100">
                {visibleRows.map((row, rowIndex) => {
                  const isSectionHeader = row[1] && !row[3] && !row[4] && !row[5] && !row[6];
                  const isCollapsed = isSectionHeader && collapsedSections.has(row[1]);
                  
                  if (isSectionHeader) {
                    return (
                      <div 
                        key={rowIndex} 
                        className="bg-slate-50 px-4 py-3 font-bold text-emerald-800 text-xs uppercase tracking-wider flex items-center justify-between cursor-pointer"
                        onClick={() => toggleSection(row[1])}
                      >
                        <span>{row[1]}</span>
                        <ChevronDown className={`w-4 h-4 transition-transform duration-200 ${isCollapsed ? '-rotate-90' : ''}`} />
                      </div>
                    );
                  }

                  const hasNoInfo = !row[3]?.trim() && !row[4]?.trim() && !(row[5] && (row[5].startsWith('http://') || row[5].startsWith('https://')));
                  
                  if (hasNoInfo) {
                    return (
                      <div key={rowIndex} className="p-4 bg-white">
                        <div className="flex items-start gap-3">
                          <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded shrink-0 mt-0.5">
                            {row[0]}
                          </span>
                          <div className="flex flex-col gap-1">
                            <h3 className="text-sm font-medium text-slate-700 leading-snug">
                              {row[1]}
                            </h3>
                            <span className="text-[11px] text-slate-400 italic">Chưa có văn bản</span>
                          </div>
                        </div>
                      </div>
                    );
                  }

                  return (
                    <div key={rowIndex} className="p-4 bg-white active:bg-slate-50 transition-colors">
                      <div className="flex items-start gap-3 mb-3">
                        <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded shrink-0 mt-0.5">
                          {row[0]}
                        </span>
                        <h3 className="text-sm font-medium text-slate-700 leading-snug">
                          {row[1]}
                        </h3>
                      </div>
                      
                      <div className="grid grid-cols-2 gap-4 mb-4 ml-8">
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] uppercase text-slate-400 font-bold tracking-widest">Số VB</span>
                          <span className="text-xs font-semibold text-slate-600">{row[3] || '---'}</span>
                        </div>
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] uppercase text-slate-400 font-bold tracking-widest">Ngày VB</span>
                          <span className="text-xs font-semibold text-slate-600">{row[4] || '---'}</span>
                        </div>
                      </div>

                      <div className="ml-8 flex items-end justify-between pt-3 border-t border-slate-50">
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] uppercase text-slate-400 font-bold tracking-widest">Cơ quan ban hành</span>
                          <span className="text-xs text-slate-500 font-medium">{row[6] || '---'}</span>
                        </div>
                        <div>
                          {row[5] && (row[5].startsWith('http://') || row[5].startsWith('https://')) ? (
                            <a 
                              href={row[5]} 
                              target="_blank" 
                              rel="noopener noreferrer"
                              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 text-indigo-700 border border-indigo-100 rounded-lg text-[10px] font-bold shadow-sm active:scale-95 transition-all"
                            >
                              <Eye className="w-3.5 h-3.5" />
                              XEM FILE
                            </a>
                          ) : (
                            <span className="text-[10px] text-slate-400 italic font-medium">
                              {row[5] || 'Không có file'}
                            </span>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
                {rows.length === 0 && (
                  <div className="p-12 text-center text-slate-400 italic">
                    Không tìm thấy dữ liệu phù hợp
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </main>

      {/* Action Loading Overlay */}
      {actionLoading && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center">
          <div className="bg-white rounded-xl shadow-xl p-8 max-w-sm w-full mx-4 flex flex-col items-center text-center">
            <div className="w-12 h-12 border-4 border-emerald-100 border-t-emerald-600 rounded-full animate-spin mb-4"></div>
            <h3 className="text-lg font-semibold text-slate-800 mb-2">Đang xử lý dữ liệu...</h3>
            <p className="text-slate-500 text-sm">
              Hệ thống đang cập nhật...
            </p>
          </div>
        </div>
      )}
      <footer className="bg-blue-900 text-white/80 py-4 mt-12">
        <div className="max-w-7xl mx-auto px-4 text-center text-xs font-medium tracking-wide">
          <p>© 2026 Công ty Điện lực Chợ Lớn</p>
        </div>
      </footer>
    </div>
  );
}
