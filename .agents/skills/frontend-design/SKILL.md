---
name: frontend-design
description: Guidelines and requirements for developing web applications, prioritizing modern, premium design aesthetics and correct technology stacks.
---

## Technology Stack
Your web applications should be built using the following technologies:
1. **Core**: Use HTML for structure and Javascript for logic.
2. **Styling (CSS)**: Use Vanilla CSS for maximum flexibility and control. Avoid using TailwindCSS unless explicitly requested.
3. **Web App**: If a more complex web app is requested, use a framework like Next.js or Vite.
4. **New Project Creation**:
   - Use `npx -y` to automatically install the script and its dependencies
   - Run the command with `--help` flag first
   - Initialize in the current directory with `./`
   - Run in non-interactive mode
5. **Running Locally**: Use `npm run dev` or equivalent dev server.

## Design Aesthetics
1. **Use Rich Aesthetics**: Create a stunning first impression with modern web design (vibrant colors, dark modes, glassmorphism, dynamic animations).
2. **Prioritize Visual Excellence**: Implement premium designs:
   - Avoid generic colors. Use curated, harmonious color palettes (e.g., HSL tailored colors, sleek dark modes).
   - Use modern typography (e.g., from Google Fonts like Inter, Roboto, or Outfit) instead of browser defaults.
   - Use smooth gradients.
   - Add subtle micro-animations for enhanced user experience.
3. **Use a Dynamic Design**: Make the interface responsive and alive with hover effects and interactive elements.
4. **Premium Designs**: Make a design that feels premium and state of the art. Avoid creating simple minimum viable products.
5. **Don't use placeholders**: If you need an image, use your generate_image tool to create a working demonstration.

## Implementation Workflow
Follow this systematic approach when building web applications:
1. **Plan and Understand**: Understand requirements, draw inspiration, outline features.
2. **Build the Foundation**: Start with `index.css` (or relevant global style file), implement core design system.
3. **Create Components**: Build components using the design system. Keep components focused and reusable.
4. **Assemble Pages**: Update the main application, ensure proper routing, implement responsive layouts.
5. **Polish and Optimize**: Review UX, ensure smooth interactions, optimize performance.

## SEO Best Practices
Automatically implement SEO best practices on every page:
- **Title Tags**: Descriptive title tags.
- **Meta Descriptions**: Compelling meta descriptions.
- **Heading Structure**: Single `<h1>` per page with proper heading hierarchy.
- **Semantic HTML**: HTML5 semantic elements.
- **Unique IDs**: Unique, descriptive IDs for browser testing.
- **Performance**: Fast page load times.

CRITICAL REMINDER: AESTHETICS ARE VERY IMPORTANT. If your web app looks simple and basic then you have FAILED!
