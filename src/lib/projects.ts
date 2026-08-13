import type { DocJSON, ExcelSheet, SlideItem, FreestyleSlide } from "./generate.functions";
import { supabase } from "./supabase";

export interface SavedProject {
  id: string;
  user_id?: string;
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

// Local storage helpers
export function getLocalProjects(): SavedProject[] {
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

export function saveLocalProjects(projects: SavedProject[]): void {
  if (typeof window === "undefined") return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(projects));
  } catch (e) {
    console.error("Failed to save projects to localStorage", e);
  }
}

// ─── SUPABASE + LOCAL STORAGE SYNC ───

/** Fetch projects from Supabase for a specific user ID, with local fallback */
export async function fetchProjectsFromSupabase(userId?: string): Promise<SavedProject[]> {
  const localProjects = getLocalProjects();
  if (!userId) return localProjects;

  try {
    const { data, error } = await supabase
      .from("projects")
      .select("*")
      .eq("user_id", userId)
      .order("updated_at", { ascending: false });

    if (error) {
      console.warn("Supabase fetch error, using local projects:", error.message);
      return localProjects;
    }

    if (data) {
      const fetched: SavedProject[] = data.map((row) => ({
        id: row.id,
        user_id: row.user_id,
        title: row.title,
        subtitle: row.subtitle || undefined,
        type: row.type as "word" | "excel" | "presentation",
        sections: row.sections || undefined,
        sheets: row.sheets || undefined,
        slides: row.slides || undefined,
        prompt: row.prompt || undefined,
        updatedAt: row.updated_at || new Date().toISOString(),
      }));

      // Cache in localStorage
      saveLocalProjects(fetched);
      return fetched;
    }
  } catch (err) {
    console.error("Error fetching projects from Supabase:", err);
  }

  return localProjects;
}

export function getProjects(): SavedProject[] {
  return getLocalProjects();
}

export function getProject(id: string): SavedProject | null {
  const projects = getLocalProjects();
  return projects.find((p) => p.id === id) || null;
}

export async function saveProject(
  project: Omit<SavedProject, "updatedAt"> & { updatedAt?: string },
  userId?: string
): Promise<SavedProject> {
  const projects = getLocalProjects();
  const updatedProject: SavedProject = {
    ...project,
    user_id: userId || project.user_id,
    updatedAt: new Date().toISOString(),
  };

  const existingIdx = projects.findIndex((p) => p.id === project.id);
  if (existingIdx > -1) {
    projects[existingIdx] = updatedProject;
  } else {
    projects.unshift(updatedProject);
  }

  // 1. Update localStorage immediately for 0ms UI reactivity
  saveLocalProjects(projects);

  // 2. Sync asynchronously to Supabase
  const effectiveUserId = userId || project.user_id;
  if (effectiveUserId) {
    try {
      const { error } = await supabase.from("projects").upsert({
        id: updatedProject.id,
        user_id: effectiveUserId,
        title: updatedProject.title,
        subtitle: updatedProject.subtitle || null,
        type: updatedProject.type,
        prompt: updatedProject.prompt || null,
        sections: updatedProject.sections || null,
        sheets: updatedProject.sheets || null,
        slides: updatedProject.slides || null,
        updated_at: updatedProject.updatedAt,
      });

      if (error) {
        console.warn("Supabase upsert error:", error.message);
      }
    } catch (err) {
      console.error("Failed to sync project to Supabase:", err);
    }
  }

  return updatedProject;
}

export async function deleteProject(id: string, userId?: string): Promise<SavedProject[]> {
  const projects = getLocalProjects();
  const filtered = projects.filter((p) => p.id !== id);
  saveLocalProjects(filtered);

  if (userId) {
    try {
      const { error } = await supabase
        .from("projects")
        .delete()
        .eq("id", id)
        .eq("user_id", userId);

      if (error) {
        console.warn("Supabase delete error:", error.message);
      }
    } catch (err) {
      console.error("Failed to delete project from Supabase:", err);
    }
  }

  return filtered;
}

export async function duplicateProject(id: string, userId?: string): Promise<SavedProject | null> {
  const target = getProject(id);
  if (!target) return null;

  const newId = crypto.randomUUID();
  const duplicated: SavedProject = {
    ...target,
    id: newId,
    user_id: userId || target.user_id,
    title: `${target.title} (نسخة)`,
    updatedAt: new Date().toISOString(),
  };

  return saveProject(duplicated, userId);
}

export async function updateProjectTitle(id: string, newTitle: string, userId?: string): Promise<SavedProject | null> {
  const target = getProject(id);
  if (!target) return null;

  const updated = {
    ...target,
    title: newTitle.trim() || target.title,
    user_id: userId || target.user_id,
    updatedAt: new Date().toISOString(),
  };

  return saveProject(updated, userId);
}
