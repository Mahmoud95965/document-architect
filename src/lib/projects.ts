import type { DocJSON, ExcelSheet, SlideItem, FreestyleSlide } from "./generate.functions";

export interface SavedProject {
  id: string;
  title: string;
  subtitle?: string;
  type: "word" | "excel" | "presentation";
  sections?: DocJSON["sections"];
  sheets?: ExcelSheet[];
  /** Generative Canvas slides (new format) — takes priority */
  slides?: FreestyleSlide[];
  /** @deprecated Legacy bullet-based slides (kept for backward compat) */
  legacySlides?: SlideItem[];
  prompt?: string;
  updatedAt: string;
}

const STORAGE_KEY = "docify_projects";

export function getProjects(): SavedProject[] {
  if (typeof window === "undefined") return [];
  try {
    const data = localStorage.getItem(STORAGE_KEY);
    if (!data) return [];
    const parsed = JSON.parse(data) as SavedProject[];
    return parsed.sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime());
  } catch (e) {
    console.error("Failed to parse projects from localStorage", e);
    return [];
  }
}

export function getProject(id: string): SavedProject | null {
  const projects = getProjects();
  return projects.find((p) => p.id === id) || null;
}

export function saveProject(project: Omit<SavedProject, "updatedAt"> & { updatedAt?: string }): SavedProject {
  const projects = getProjects();
  const updatedProject: SavedProject = {
    ...project,
    updatedAt: new Date().toISOString(),
  };

  const existingIdx = projects.findIndex((p) => p.id === project.id);
  if (existingIdx > -1) {
    projects[existingIdx] = updatedProject;
  } else {
    projects.push(updatedProject);
  }

  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(projects));
  } catch (e) {
    console.error("Failed to save projects to localStorage", e);
  }
  return updatedProject;
}

export function deleteProject(id: string): SavedProject[] {
  const projects = getProjects();
  const filtered = projects.filter((p) => p.id !== id);
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(filtered));
  } catch (e) {
    console.error("Failed to delete project from localStorage", e);
  }
  return filtered;
}
