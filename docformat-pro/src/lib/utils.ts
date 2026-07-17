import { clsx, type ClassValue } from "clsx";

export function cn(...inputs: ClassValue[]) {
  return clsx(inputs).split(' ').filter(Boolean).join(' ').trim() || undefined;
}
