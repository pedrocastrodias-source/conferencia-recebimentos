import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://neypkxgsvikorzjymemh.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im5leXBreGdzdmlrb3J6anltZW1oIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjYwMDE1NjksImV4cCI6MjA4MTU3NzU2OX0.ApkZFrUXP9r7qXrZU-RyazdI06s1DXxe_kZs5OQRCfk';

export const supabase = createClient(supabaseUrl, supabaseKey);
