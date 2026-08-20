import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  base: "/channel-review-workbench/",
  plugins: [react()],
});
